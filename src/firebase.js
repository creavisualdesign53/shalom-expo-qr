
import { initializeApp } from 'firebase/app'
import { getAuth } from 'firebase/auth'
import {
  initializeFirestore,
  persistentLocalCache,
  persistentMultipleTabManager
} from 'firebase/firestore'

export const firebaseConfig = {
  apiKey: "AIzaSyAbHb4WWw9U7ptKbzUfVvQXC1CVTurC8sY",
  authDomain: "shalom-expo-2026-a6aae.firebaseapp.com",
  projectId: "shalom-expo-2026-a6aae",
  storageBucket: "shalom-expo-2026-a6aae.firebasestorage.app",
  messagingSenderId: "614526590692",
  appId: "1:614526590692:web:2223282ac8a26e54fbd6f1"
}

const app = initializeApp(firebaseConfig)

export const auth = getAuth(app)
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({
    tabManager: persistentMultipleTabManager()
  })
})