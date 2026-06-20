// ════════════════════════════════════════════════════════════════
// FeeTrack UG — Supabase Connection Test
// ════════════════════════════════════════════════════════════════
// A standalone test page, completely separate from the main app, whose
// only purpose is to prove the database connection works before we touch
// any real app code. It tries to read the `schools` table and shows
// exactly what comes back — including the real error message if
// anything fails, rather than hiding it.
//
// HOW TO USE THIS FILE:
// This is meant to be deployed (e.g. via Vercel/Netlify) alongside the
// main app, NOT run inside the Claude chat preview — the Supabase
// library isn't available in that preview environment. Once the project
// is hosted somewhere real, open this page first to confirm the database
// connection works, before relying on the main app's database features.
// ════════════════════════════════════════════════════════════════

import { useState, useEffect } from "react";
import { supabase } from "./supabaseClient";

export default function ConnectionTest() {
  const [status, setStatus] = useState("checking"); // checking | success | error
  const [details, setDetails] = useState("");
  const [schools, setSchools] = useState([]);

  useEffect(() => {
    async function testConnection() {
      try {
        const { data, error } = await supabase.from("schools").select("*");
        if (error) {
          setStatus("error");
          setDetails(error.message);
          return;
        }
        setStatus("success");
        setSchools(data || []);
        setDetails(`Connected successfully. The schools table currently has ${data.length} row(s).`);
      } catch (err) {
        setStatus("error");
        setDetails(err.message || "Unknown error — could not reach Supabase at all.");
      }
    }
    testConnection();
  }, []);

  const box = {
    maxWidth: 480,
    margin: "60px auto",
    padding: 24,
    borderRadius: 14,
    fontFamily: "system-ui, sans-serif",
    border: "1px solid #e2e8f0",
    background: status === "error" ? "#fef2f2" : status === "success" ? "#f0fdf4" : "#f8fafc",
  };

  return (
    <div style={box}>
      <h2 style={{ marginTop: 0 }}>FeeTrack UG — Database Connection Test</h2>
      {status === "checking" && <p>Checking connection to Supabase…</p>}
      {status === "success" && (
        <>
          <p style={{ color: "#15803d", fontWeight: 700 }}>✓ Connected successfully</p>
          <p>{details}</p>
          {schools.length > 0 && (
            <pre style={{ background: "#fff", padding: 12, borderRadius: 8, overflowX: "auto", fontSize: 12 }}>
              {JSON.stringify(schools, null, 2)}
            </pre>
          )}
        </>
      )}
      {status === "error" && (
        <>
          <p style={{ color: "#b91c1c", fontWeight: 700 }}>✗ Connection failed</p>
          <p>Error message from Supabase:</p>
          <pre style={{ background: "#fff", padding: 12, borderRadius: 8, overflowX: "auto", fontSize: 12, color: "#b91c1c" }}>
            {details}
          </pre>
        </>
      )}
    </div>
  );
}
