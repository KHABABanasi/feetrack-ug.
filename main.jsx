import React from "react";
import { createRoot } from "react-dom/client";
import App from "./School-fee-manager.jsx";
import ConnectionTest from "./ConnectionTest.jsx";

// Visit yoursite.com/?test to see the database connection test page instead
// of the main app — useful for confirming Supabase is wired up correctly
// without needing to log into the real app first.
const showTestPage = new URLSearchParams(window.location.search).has("test");

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    {showTestPage ? <ConnectionTest /> : <App />}
  </React.StrictMode>
);
