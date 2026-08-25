import { initializeApp, getApps, getApp } from 'firebase/app';
import { initializeFirestore, getFirestore, memoryLocalCache } from 'firebase/firestore';
import { getAuth } from 'firebase/auth';

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || "AIzaSyDivaTH9k1C7y_2KDNvtLvCGELV4OaTECs",
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || "zeta-walker-3n50x.firebaseapp.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || "zeta-walker-3n50x",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || "zeta-walker-3n50x.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || "209095575347",
  appId: import.meta.env.VITE_FIREBASE_APP_ID || "1:209095575347:web:e0cc7a7c778008393c2046"
};

const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();

// Using memoryLocalCache with force long polling to avoid IndexedDB lock closure errors in preview sandboxes
let firestoreDb: any;
try {
  firestoreDb = initializeFirestore(app, {
    localCache: memoryLocalCache(),
    experimentalForceLongPolling: true,
  });
} catch {
  try {
    firestoreDb = getFirestore(app);
  } catch (e) {
    console.warn('Firestore fallback warning:', e);
  }
}

export const db = firestoreDb;
export const auth = getAuth(app);

