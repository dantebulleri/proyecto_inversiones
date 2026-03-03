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
    apiKey: "AIzaSyCicouXsj6c53UBLDaPjus9dnGP2QOXnR0",
    authDomain: "moto-flip.firebaseapp.com",
    projectId: "moto-flip",
    storageBucket: "moto-flip.firebasestorage.app",
    messagingSenderId: "260447843449",
    appId: "1:260447843449:web:6753de1e3b34a5cbaf50a8"
};

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();

// Soporte offline
db.enablePersistence({ synchronizeTabs: true }).catch(() => {});

// Idioma en espanol para emails de Firebase
auth.languageCode = 'es';
