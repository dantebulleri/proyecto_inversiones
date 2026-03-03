/* ============================================================
   FIREBASE CONFIG - MotoFlip
   ============================================================
   INSTRUCCIONES para configurar Firebase:

   1. Anda a https://console.firebase.google.com
   2. Crea un proyecto nuevo (ej: "motoflip")
   3. En el proyecto, activa Authentication:
      - Authentication > Sign-in method > Email/Password > Habilitar
   4. Activa Firestore Database:
      - Firestore Database > Crear base de datos > Modo de prueba
   5. En Configuracion del proyecto (engranaje) > General:
      - Baja hasta "Tus apps" > click en </> (Web)
      - Registra la app (nombre: "MotoFlip")
      - Copia los valores del objeto firebaseConfig aqui abajo
   ============================================================ */

const firebaseConfig = {
    apiKey: "TU_API_KEY",
    authDomain: "TU_PROYECTO.firebaseapp.com",
    projectId: "TU_PROYECTO",
    storageBucket: "TU_PROYECTO.firebasestorage.app",
    messagingSenderId: "TU_SENDER_ID",
    appId: "TU_APP_ID"
};

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();

// Soporte offline
db.enablePersistence({ synchronizeTabs: true }).catch(() => {});

// Idioma en espanol para emails de Firebase
auth.languageCode = 'es';
