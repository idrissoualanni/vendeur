/**
 * =================================================================
 *  BACKEND GOOGLE APPS SCRIPT - GESTION AVANCÉE DE VENDEURS
 * =================================================================
 *  - Gestion des Rôles (Admin, Informaticien, Représentant)
 *  - Authentification des utilisateurs
 *  - Structure de données normalisée sur plusieurs feuilles
 *  - Gestion des produits simples et variables
 *  - Upload d'images vers Google Drive
 */

// --- CONFIGURATION DES FEUILLES ---
const SHEETS = {
  USERS: 'Utilisateurs',
  VENDORS: 'Vendeurs',
  PRODUCTS: 'Produits',
  VARIATIONS: 'Produits_Variations'
};

const ROLES = {
  ADMIN: 'Administrateur',
  IT: 'Informaticien',
  REP: 'Représentant'
};

// --- FONCTION D'INITIALISATION (À EXÉCUTER UNE FOIS MANUELLEMENT) ---
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Création des feuilles
  Object.values(SHEETS).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });

  // 2. Définition des en-têtes
  const headers = {
    [SHEETS.USERS]: ['Email', 'Password', 'Role', 'VendeurID', 'Token'],
    [SHEETS.VENDORS]: ['VendeurID', 'NomVendeur', 'Pays', 'WooCommerceInfo', 'CrééParEmail'],
    [SHEETS.PRODUCTS]: ['ProduitID', 'VendeurID', 'NomProduit', 'SKU', 'Type', 'Catégories', 'ImageURL', 'StatutRécupération', 'Attributs'],
    [SHEETS.VARIATIONS]: ['VariationID', 'ProduitID', 'CombinaisonAttributs', 'Prix', 'SKU_Variation']
  };

  for (const sheetName in headers) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers[sheetName]);
    }
  }

  // 3. Création d'un utilisateur Admin par défaut
  const userSheet = ss.getSheetByName(SHEETS.USERS);
  if (userSheet.getLastRow() < 2) {
    // IMPORTANT: Changez ce mot de passe !
    userSheet.appendRow(['admin@example.com', 'admin123', ROLES.ADMIN, 'ADMIN_ID', '']);
    SpreadsheetApp.getUi().alert('Configuration terminée. Un utilisateur "admin@example.com" avec le mot de passe "admin123" a été créé. Pensez à le changer !');
  } else {
    SpreadsheetApp.getUi().alert('La configuration semble déjà avoir été effectuée.');
  }
}


// --- GESTION DES REQUÊTES HTTP (API) ---

function doGet(e) {
  // Le doGet peut être utilisé pour récupérer des données publiquement ou après authentification
  // Par simplicité, on garde la logique principale dans doPost et après authentification
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'API en ligne. Utilisez POST pour les actions.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let response;
  try {
    const request = JSON.parse(e.postData.contents);
    const { action, payload, token } = request;
    
    // L'action 'login' est la seule autorisée sans token
    if (action === 'login') {
      response = loginUser(payload.email, payload.password);
    } else {
      // Toutes les autres actions nécessitent une authentification
      const user = getAuthenticatedUser(token);
      if (!user) throw new Error("Accès non autorisé. Token invalide ou expiré.");

      switch (action) {
        case 'getDashboardData':
          response = getDashboardData(user);
          break;
        case 'registerRepresentative':
          // Seuls Admin et IT peuvent enregistrer un représentant
          if (user.role !== ROLES.ADMIN && user.role !== ROLES.IT) {
            throw new Error("Permission refusée.");
          }
          response = registerRepresentative(payload, user.email);
          break;
        case 'addProduct':
          response = addProduct(payload, user);
          break;
        // Ajoutez d'autres actions ici (ex: 'updateProductStatus', 'getVendorsList', etc.)
        default:
          throw new Error(`Action "${action}" non reconnue.`);
      }
    }
  } catch (error) {
    response = { status: 'error', message: error.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}


// --- LOGIQUE D'AUTHENTIFICATION ---

function loginUser(email, password) {
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const data = userSheet.getDataRange().getValues();
  
  // On cherche l'utilisateur (on ignore la ligne d'en-tête)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase() && data[i][1] === password) {
      const token = `TOKEN_${Utilities.getUuid()}`;
      // Stocker le token dans la feuille (colonne 5, index 4)
      userSheet.getRange(i + 1, 5).setValue(token);
      
      return {
        status: 'success',
        token: token,
        user: {
          email: data[i][0],
          role: data[i][2],
          vendeurId: data[i][3]
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
    if (data[i][4] === token) { // Colonne du Token
      return { email: data[i][0], role: data[i][2], vendeurId: data[i][3] };
    }
  }
  return null;
}


// --- LOGIQUE MÉTIER (ACTIONS DE L'API) ---

function getDashboardData(user) {
  const products = getSheetDataAsObject(SHEETS.PRODUCTS);
  const vendors = getSheetDataAsObject(SHEETS.VENDORS);
  
  if (user.role === ROLES.REP) {
    // Un représentant ne voit que ses produits et son info vendeur
    const myVendor = vendors.filter(v => v.VendeurID === user.vendeurId);
    const myProducts = products.filter(p => p.VendeurID === user.vendeurId);
    return { status: 'success', vendors: myVendor, products: myProducts };
  } else {
    // Admin et IT voient tout
    return { status: 'success', vendors: vendors, products: products };
  }
}

function registerRepresentative(payload, createdByEmail) {
  const { nomVendeur, pays, email, password } = payload;
  if (!nomVendeur || !pays || !email || !password) {
    throw new Error("Informations manquantes pour l'inscription.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
    const vendorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VENDORS);

    // Vérifier si l'email existe déjà
    const emails = userSheet.getRange(2, 1, userSheet.getLastRow(), 1).getValues().flat();
    if (emails.includes(email)) {
      throw new Error(`L'email "${email}" est déjà utilisé.`);
    }

    const newVendeurID = `VEN_${Utilities.getUuid()}`;
    
    // 1. Ajouter le vendeur
    vendorSheet.appendRow([newVendeurID, nomVendeur, pays, '', createdByEmail]);
    
    // 2. Ajouter l'utilisateur représentant
    userSheet.appendRow([email, password, ROLES.REP, newVendeurID, '']);
    
    return { status: 'success', message: `Le représentant ${nomVendeur} a été créé avec succès.` };
  } finally {
    lock.releaseLock();
  }
}

function addProduct(payload, user) {
  const { nomProduit, sku, type, categories, attributs, variations, imageFile } = payload;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const productSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
    const variationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VARIATIONS);
    
    let vendeurId = user.vendeurId;
    // Si c'est un Admin/IT qui ajoute, il doit spécifier pour qui
    if (user.role !== ROLES.REP) {
      if (!payload.vendeurId) throw new Error("Un vendeur doit être spécifié.");
      vendeurId = payload.vendeurId;
    }
    
    // Gérer l'upload d'image
    let imageUrl = '';
    if (imageFile && imageFile.base64) {
      imageUrl = uploadFileToDrive(imageFile.base64, imageFile.name, imageFile.type);
    }
    
    const newProduitID = `PROD_${Utilities.getUuid()}`;
    
    // Ajouter le produit principal
    productSheet.appendRow([
      newProduitID,
      vendeurId,
      nomProduit,
      sku,
      type,
      categories, // ex: "T-shirt, Homme"
      imageUrl,
      'Non récupéré',
      JSON.stringify(attributs) // ex: [{"nom":"Couleur", "valeurs":"Bleu, Vert, Rouge"}, {"nom":"Taille", "valeurs":"S, M, L"}]
    ]);
    
    // Si le produit est variable, ajouter les variations
    if (type === 'variable' && variations && variations.length > 0) {
      variations.forEach(v => {
        const newVariationID = `VAR_${Utilities.getUuid()}`;
        variationSheet.appendRow([
          newVariationID,
          newProduitID,
          JSON.stringify(v.combinaison), // ex: {"Couleur": "Bleu", "Taille": "M"}
          v.prix,
          v.sku_variation
        ]);
      });
    }
    
    return { status: 'success', message: `Produit "${nomProduit}" ajouté.` };
    
  } finally {
    lock.releaseLock();
  }
}


// --- FONCTIONS UTILITAIRES ---

function uploadFileToDrive(base64Data, fileName, mimeType) {
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, `img_${new Date().getTime()}_${fileName}`);
    
    // Optionnel: Spécifier un dossier dans Google Drive
    // const folder = DriveApp.getFolderById("ID_DE_VOTRE_DOSSIER_DRIVE");
    // const file = folder.createFile(blob);
    const file = DriveApp.getRootFolder().createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
}

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