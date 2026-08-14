import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { MsalProvider } from "@azure/msal-react";
import { msalInstance } from "./authConfig";
import './index.css'
import './dashboard.css'
import './studio.css'
import App from './App.jsx'

// Ensure MSAL is initialized and redirects are processed before the app renders
msalInstance.initialize().then(() => {
  return msalInstance.handleRedirectPromise();
}).then(() => {
  createRoot(document.getElementById('root')).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>,
  );
}).catch(err => {
    console.error("MSAL Redirect Processing Error:", err);
});
