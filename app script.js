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
// Noms des feuilles de calcul qui serviront de base de données.
const SHEETS = {
  USERS: 'Utilisateurs',
  PRODUCTS: 'Produits'
};

// Rôles des utilisateurs pour gérer les permissions.
const ROLES = {
  ADMIN: 'Admin',
  SELLER: 'Vendeur'
};

/**
 * =================================================================
 * INITIALISATION (À EXÉCUTER UNE SEULE FOIS MANUELLEMENT)
 * =================================================================
 * Cette fonction prépare la feuille de calcul Google Sheets.
 * 1. Elle crée les feuilles 'Utilisateurs' et 'Produits' si elles n'existent pas.
 * 2. Elle ajoute les en-têtes de colonnes.
 * 3. Elle crée un utilisateur 'Admin' par défaut.
 * Pour l'exécuter : Dans l'éditeur Apps Script, sélectionnez 'initialSetup' dans le menu déroulant et cliquez sur 'Exécuter'.
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Créer les feuilles
  Object.values(SHEETS).forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });

  const userSheet = ss.getSheetByName(SHEETS.USERS);
  const productSheet = ss.getSheetByName(SHEETS.PRODUCTS);

  // Ajouter les en-têtes si les feuilles sont vides
  if (userSheet.getLastRow() === 0) {
    userSheet.appendRow(['UserID', 'Email', 'Password', 'Role', 'Token']);
  }
  if (productSheet.getLastRow() === 0) {
    productSheet.appendRow(['ProductID', 'VendeurID', 'NomProduit', 'Prix', 'ImageURL']);
  }

  // Créer un admin par défaut s'il n'y en a pas
  if (userSheet.getLastRow() < 2) {
    const adminEmail = 'admin@example.com';
    const adminPassword = 'adminpassword'; // IMPORTANT: Changez ce mot de passe !
    userSheet.appendRow([`user_${Utilities.getUuid()}`, adminEmail, adminPassword, ROLES.ADMIN, '']);
    
    // Affiche une alerte pour informer l'utilisateur.
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Configuration terminée !',
      `Un utilisateur administrateur a été créé avec l'email "${adminEmail}" et le mot de passe "${adminPassword}". N'oubliez pas de changer le mot de passe.`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * =================================================================
 * GESTION DES REQUÊTES HTTP (API)
 * =================================================================
 * C'est le point d'entrée principal de notre API.
 * Toutes les requêtes de l'application front-end arriveront ici.
 */
function doPost(e) {
  let response;
  try {
    const request = JSON.parse(e.postData.contents);
    const { action, payload, token } = request;

    // Les actions 'login' et 'registerRepresentative' sont autorisées sans token
    if (action === 'login') {
      response = loginUser(payload.email, payload.password);
    } else if (action === 'registerRepresentative') {
      // On rend cette action publique pour l'auto-inscription des vendeurs
      response = registerRepresentative(payload, 'Self-Registered'); // L'email du créateur n'est pas pertinent ici
    }
    else {
      // Toutes les autres actions nécessitent une authentification
      const user = getAuthenticatedUser(token);
      if (!user) throw new Error("Accès non autorisé. Token invalide ou expiré.");

      switch (action) {
        case 'getDashboardData':
          response = getDashboardData(user);
          break;
        // La logique de 'registerRepresentative' a été déplacée pour être publique
        case 'addProduct':
          response = addProduct(payload, user);
          break;
        // Ajoutez d'autres actions ici (ex: 'updateProductStatus', 'getVendorsList', etc.)
        default:
          throw new Error(`Action "${action}" non reconnue.`);
      }
    }
  } catch (error) {
    // Gestion centralisée des erreurs
    response = { status: 'error', message: error.message };
  }

  // On retourne la réponse au format JSON.
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Gère les requêtes "preflight" envoyées par les navigateurs pour la vérification CORS.
 * C'est essentiel pour permettre à une application web externe (comme celle sur Vercel)
 * d'appeler cette API.
 */
function doOptions(e) {
  return ContentService.createTextOutput()
    .setHeader('Access-Control-Allow-Origin', '*') // Permet à n'importe quel domaine d'appeler l'API
    .setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * =================================================================
 * FONCTIONS D'AUTHENTIFICATION
 * =================================================================
 */

/**
 * Inscrit un nouveau vendeur.
 * @param {object} payload - Doit contenir {email, password}.
 */
function signup(payload) {
  const { email, password } = payload;
  if (!email || !password) {
    throw new Error("L'email et le mot de passe sont requis pour l'inscription.");
  }

  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const emails = userSheet.getRange(2, 2, userSheet.getLastRow(), 1).getValues().flat();

  // Vérifier si l'email est déjà utilisé
  if (emails.includes(email)) {
    throw new Error(`L'email "${email}" est déjà utilisé.`);
  }

  // Ajouter le nouvel utilisateur avec le rôle 'Vendeur'
  userSheet.appendRow([`user_${Utilities.getUuid()}`, email, password, ROLES.SELLER, '']);

  return { status: 'success', message: 'Vendeur inscrit avec succès. Vous pouvez maintenant vous connecter.' };
}

/**
 * Connecte un utilisateur (Admin ou Vendeur).
 * @param {object} payload - Doit contenir {email, password}.
 */
function login(payload) {
  const { email, password } = payload;
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const data = userSheet.getDataRange().getValues();

  // On cherche l'utilisateur dans la feuille (on ignore les en-têtes)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // row[1] est l'email, row[2] est le mot de passe
    if (row[1].toLowerCase() === email.toLowerCase() && row[2] === password) {
      const token = `token_${Utilities.getUuid()}`;
      // On stocke le nouveau token dans la feuille
      userSheet.getRange(i + 1, 5).setValue(token); // Colonne 'Token'

      return {
        status: 'success',
        token: token,
        user: {
          userId: row[0], // UserID
          email: row[1], // Email
          role: row[3]   // Role
        }
      };
    }
  }

  // Si on ne trouve pas l'utilisateur
  throw new Error("Email ou mot de passe incorrect.");
}

/**
 * Vérifie si un token est valide et retourne les informations de l'utilisateur.
 * @param {string} token - Le token d'authentification.
 * @returns {object|null} - L'objet utilisateur ou null.
 */
function getAuthenticatedUser(token) {
  if (!token) return null;
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  const data = userSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[4] === token) { // Colonne 'Token'
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
 * LOGIQUE MÉTIER (ACTIONS DE L'API)
 * =================================================================
 */

/**
 * Ajoute un nouveau produit pour le vendeur connecté.
 * @param {object} payload - Contient les infos du produit {nom, prix, imageFile}.
 * @param {object} user - L'objet utilisateur authentifié.
 */
function addProduct(payload, user) {
  const { nom, prix, imageFile } = payload;
  if (!nom || !prix) {
    throw new Error("Le nom et le prix du produit sont requis.");
  }

  let imageUrl = '';
  // Si une image est fournie, on l'upload sur Google Drive
  if (imageFile && imageFile.base64) {
    imageUrl = uploadFileToDrive(imageFile.base64, imageFile.name, imageFile.type);
  }

  const productSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCTS);
  productSheet.appendRow([
    `prod_${Utilities.getUuid()}`, // ProductID
    user.userId,                  // VendeurID
    nom,
    prix,
    imageUrl
  ]);

  return { status: 'success', message: `Produit "${nom}" ajouté avec succès.` };
}

/**
 * Récupère la liste des produits pour le vendeur connecté.
 * @param {object} user - L'objet utilisateur authentifié.
 */
function getSellerProducts(user) {
  const allProducts = getSheetDataAsObject(SHEETS.PRODUCTS);
  const sellerProducts = allProducts.filter(p => p.VendeurID === user.userId);

  return { status: 'success', products: sellerProducts };
}

/**
 * Récupère toutes les données pour le tableau de bord de l'administrateur.
 */
function getAdminDashboard() {
  const users = getSheetDataAsObject(SHEETS.USERS);
  const products = getSheetDataAsObject(SHEETS.PRODUCTS);

  // On retire les mots de passe et tokens avant de les renvoyer
  const sanitizedUsers = users.map(u => ({
    UserID: u.UserID,
    Email: u.Email,
    Role: u.Role
  }));

  return { status: 'success', users: sanitizedUsers, products: products };
}


/**
 * =================================================================
 * FONCTIONS UTILITAIRES
 * =================================================================
 */

/**
 * Convertit les données d'une feuille en un tableau d'objets.
 * @param {string} sheetName - Le nom de la feuille.
 * @returns {Array<object>}
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

/**
 * Uploade un fichier sur Google Drive et retourne son URL publique.
 * @param {string} base64Data - Les données du fichier en base64.
 * @param {string} fileName - Le nom du fichier.
 * @param {string} mimeType - Le type MIME du fichier.
 */
function uploadFileToDrive(base64Data, fileName, mimeType) {
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, `img_${new Date().getTime()}_${fileName}`);
  
  // On crée le fichier à la racine du Drive
  const file = DriveApp.getRootFolder().createFile(blob);

  // On rend le fichier visible par tous ceux qui ont le lien
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}
