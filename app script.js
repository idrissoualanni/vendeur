/**
 * =================================================================
 *  BACKEND GOOGLE APPS SCRIPT - GESTION SIMPLE DE VENDEURS
 * =================================================================
 *  Ce script met en place une API simple pour :
 *  - L'inscription et la connexion des vendeurs.
 *  - L'ajout de produits par les vendeurs.
 *  - Un tableau de bord pour l'administrateur pour voir tous les utilisateurs et produits.
 */

// --- CONFIGURATION ---
const SHEETS = {
  USERS: 'Utilisateurs',
  PRODUCTS: 'Produits',
  VENDEURS: 'Vendeurs'
};

const ROLES = {
  ADMIN: 'Administrateur',
  IT: 'Informaticien',
  SELLER: 'Représentant'
};

const ALLOWED_ORIGIN = '*'; // À restreindre en production

/**
 * =================================================================
 * INITIALISATION (À EXÉCUTER UNE SEULE FOIS)
 * =================================================================
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.values(SHEETS).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });

  const userSheet = ss.getSheetByName(SHEETS.USERS);
  const productSheet = ss.getSheetByName(SHEETS.PRODUCTS);
  const vendeurSheet = ss.getSheetByName(SHEETS.VENDEURS);

  if (userSheet.getLastRow() === 0) {
    userSheet.appendRow(['UserID', 'Email', 'Password', 'Role', 'Token']);
  }
  if (productSheet.getLastRow() === 0) {
    productSheet.appendRow(['ProductID', 'VendeurID', 'NomProduit', 'SKU', 'Type', 'Categories', 'ImageURL', 'StatutRécupération']);
  }
  if (vendeurSheet.getLastRow() === 0) {
    vendeurSheet.appendRow(['VendeurID', 'NomVendeur', 'Pays', 'CrééParEmail', 'Email']);
  }

  if (userSheet.getLastRow() < 2) {
    const adminEmail = 'admin@example.com';
    const adminPassword = 'adminpassword';
    userSheet.appendRow([`user_${Utilities.getUuid()}`, adminEmail, adminPassword, ROLES.ADMIN, '']);
    
    SpreadsheetApp.getUi().alert(
      'Configuration terminée !',
      `Admin créé : "${adminEmail}" / "${adminPassword}"`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * =================================================================
 * GESTION DES REQUÊTES HTTP (API)
 * =================================================================
 */

function doOptions(e) {
  return createCORSResponse('');
}

function doGet(e) {
  const response = { status: 'success', message: 'API en ligne. Utilisez POST pour les actions.' };
  return createCORSResponse(JSON.stringify(response));
}

function doPost(e) {
  let response;
  try {
    const request = JSON.parse(e.postData.contents);
    const { action, payload, token } = request;

    if (action === 'login') {
      response = login(payload);
    } else if (action === 'signup') {
      response = signup(payload);
    } else if (action === 'registerRepresentative') {
      response = registerRepresentative(payload);
    } else {
      const user = getAuthenticatedUser(token);
      if (!user) {
        throw new Error("Accès non autorisé. Token invalide ou expiré.");
      }

      switch (action) {
        case 'getDashboardData':
          if (user.role !== ROLES.ADMIN && user.role !== ROLES.IT) {
            throw new Error("Accès refusé. Réservé aux administrateurs.");
          }
          response = getAdminDashboard();
          break;
        case 'addProduct':
          response = addProduct(payload, user);
          break;
        case 'getSellerProducts':
          response = getSellerProducts(user);
          break;
        default:
          throw new Error(`Action "${action}" non reconnue.`);
      }
    }
  } catch (error) {
    response = { status: 'error', message: error.message };
  }

  return createCORSResponse(JSON.stringify(response));
}

function createCORSResponse(content) {
  return ContentService.createTextOutput(content)
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', ALLOWED_ORIGIN)
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * =================================================================
 * FONCTIONS D'AUTHENTIFICATION
 * =================================================================
 */

function signup(payload) {
  const { email, password } = payload;
  if (!email || !password) {
    throw new Error("L'email et le mot de passe sont requis.");
  }

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const emails = userSheet.getRange(2, 2, userSheet.getLastRow(), 1).getValues().flat();

  if (emails.includes(email)) {
    throw new Error(`L'email "${email}" est déjà utilisé.`);
  }

  userSheet.appendRow([`user_${Utilities.getUuid()}`, email, password, ROLES.SELLER, '']);

  return { status: 'success', message: 'Compte créé avec succès. Vous pouvez maintenant vous connecter.' };
}

function registerRepresentative(payload) {
  const { nomVendeur, pays, email, password } = payload;
  if (!email || !password || !nomVendeur || !pays) {
    throw new Error("Tous les champs sont requis.");
  }

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const vendeurSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VENDEURS);
  
  const emails = userSheet.getRange(2, 2, userSheet.getLastRow(), 1).getValues().flat();
  if (emails.includes(email)) {
    throw new Error(`L'email "${email}" est déjà utilisé.`);
  }

  const vendeurId = `vendeur_${Utilities.getUuid()}`;
  const userId = `user_${Utilities.getUuid()}`;

  // Créer le vendeur
  vendeurSheet.appendRow([vendeurId, nomVendeur, pays, email, email]);
  
  // Créer l'utilisateur
  userSheet.appendRow([userId, email, password, ROLES.SELLER, '']);

  return { status: 'success', message: 'Compte vendeur créé avec succès.' };
}

function login(payload) {
  const { email, password } = payload;
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const data = userSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1].toLowerCase() === email.toLowerCase() && row[2] === password) {
      const token = `token_${Utilities.getUuid()}`;
      userSheet.getRange(i + 1, 5).setValue(token);

      return {
        status: 'success',
        token: token,
        user: {
          userId: row[0],
          email: row[1],
          role: row[3]
        }
      };
    }
  }

  throw new Error("Email ou mot de passe incorrect.");
}

function getAuthenticatedUser(token) {
  if (!token) return null;
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const data = userSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[4] === token) {
      return {
        userId: row[0],
        email: row[1],
        role: row[3]
      };
    }
  }
  return null;
}

/**
 * =================================================================
 * LOGIQUE MÉTIER
 * =================================================================
 */

function addProduct(payload, user) {
  const { nom, prix, sku, type, categories, imageFile, vendeurId } = payload;
  if (!nom || !type) {
    throw new Error("Le nom et le type du produit sont requis.");
  }

  let imageUrl = '';
  if (imageFile && imageFile.base64) {
    imageUrl = uploadFileToDrive(imageFile.base64, imageFile.name, imageFile.type);
  }

  const productSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
  const targetVendeurId = vendeurId || user.userId;

  productSheet.appendRow([
    `prod_${Utilities.getUuid()}`,
    targetVendeurId,
    nom,
    sku || '',
    type,
    categories || '',
    imageUrl,
    'En attente'
  ]);

  return { status: 'success', message: `Produit "${nom}" ajouté avec succès.` };
}

function getSellerProducts(user) {
  const allProducts = getSheetDataAsObject(SHEETS.PRODUCTS);
  const sellerProducts = allProducts.filter(p => p.VendeurID === user.userId);

  return { status: 'success', products: sellerProducts };
}

function getAdminDashboard() {
  const users = getSheetDataAsObject(SHEETS.USERS);
  const products = getSheetDataAsObject(SHEETS.PRODUCTS);
  const vendeurs = getSheetDataAsObject(SHEETS.VENDEURS);

  const sanitizedUsers = users.map(u => ({
    UserID: u.UserID,
    Email: u.Email,
    Role: u.Role
  }));

  return { 
    status: 'success', 
    users: sanitizedUsers, 
    products: products,
    vendors: vendeurs
  };
}

/**
 * =================================================================
 * FONCTIONS UTILITAIRES
 * =================================================================
 */

function getSheetDataAsObject(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

function uploadFileToDrive(base64Data, fileName, mimeType) {
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, `img_${new Date().getTime()}_${fileName}`);
  
  const file = DriveApp.getRootFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}
