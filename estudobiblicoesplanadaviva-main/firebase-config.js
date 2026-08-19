import { initializeApp } from "https://www.gstatic.com/firebasejs/12.11.0/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/12.11.0/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/12.11.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyAvv9_G2-qSrwADgoQanPdkpBRLDr-d4ps",
  authDomain: "estudosbiblicosesplanadaviva.firebaseapp.com",
  projectId: "estudosbiblicosesplanadaviva",
  storageBucket: "estudosbiblicosesplanadaviva.firebasestorage.app",
  messagingSenderId: "831476101977",
  appId: "1:831476101977:web:ee99fa8934878c15fc9248",
  measurementId: "G-GRH2VHH519"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

export { app, auth, db };