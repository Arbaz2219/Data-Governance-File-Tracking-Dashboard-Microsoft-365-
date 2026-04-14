import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { MsalProvider } from "@azure/msal-react";
import { msalInstance } from "./authConfig";
import './index.css'
import App from './App.jsx'

// Ensure MSAL is initialized before the app renders
msalInstance.initialize().then(() => {
  createRoot(document.getElementById('root')).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>,
  );
});
