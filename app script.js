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
 * GESTION DES REQUÊTES HTTP (API) - CORRIGÉ POUR CORS
 * =================================================================
 * Ce bloc gère toutes les requêtes entrantes et inclut la logique
 * pour résoudre les erreurs CORS entre Vercel et Google Apps Script.
 */

// L'URL de votre application frontend. Seules les requêtes provenant de ce domaine seront autorisées.
const ALLOWED_ORIGIN = 'https://vendeur-theta.vercel.app';

/**
 * Gère les requêtes "preflight" (OPTIONS) du navigateur.
 * Le navigateur envoie une requête OPTIONS avant la requête POST réelle pour vérifier
 * si le serveur autorise la communication. C'est la première étape de la validation CORS.
 */
function doOptions(e) {
  return ContentService.createTextOutput()
    .setHeader('Access-Control-Allow-Origin', ALLOWED_ORIGIN)
    .setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Gère les requêtes GET. Utile pour vérifier si l'API est en ligne.
 * Chaque réponse GET inclut aussi les en-têtes CORS.
 */
function doGet(e) {
    const response = { status: 'success', message: 'API en ligne. Utilisez POST pour les actions.' };
    return ContentService.createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin', ALLOWED_ORIGIN)
        .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
 * Gère les requêtes POST contenant les données (ex: login, addProduct).
 * C'est le point d'entrée principal de l'API.
 * IMPORTANT : La réponse à cette requête DOIT également contenir les en-têtes CORS
 * pour que le navigateur autorise le code JavaScript à lire le résultat.
 */
function doPost(e) {
  let response;
  try {
    const request = JSON.parse(e.postData.contents);
    const { action, payload, token } = request;

    // Les actions 'login' et 'signup' sont autorisées sans token
    if (action === 'login') {
      response = login(payload);
    } else if (action === 'signup') {
      response = signup(payload);
    }
    else {
      // Toutes les autres actions nécessitent une authentification
      const user = getAuthenticatedUser(token);
      if (!user) throw new Error("Accès non autorisé. Token invalide ou expiré.");

      switch (action) {
        case 'getDashboardData': // Devrait être sécurisé
           if(user.role !== ROLES.ADMIN) throw new Error("Accès refusé.");
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
    // Gestion centralisée des erreurs
    response = { status: 'error', message: error.message };
  }

  // On retourne la réponse au format JSON avec les en-têtes CORS.
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', ALLOWED_ORIGIN)
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-control-Allow-Headers', 'Content-Type');
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
