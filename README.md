# FeeTrack UG

A school fee management system for Ugandan schools, built with React + Vite, connected to a Supabase (Postgres) database.

## Local setup (only needed if running this on your own computer)

1. Copy `.env.example` to a new file called `.env`, and fill in your own Supabase project's URL and key (found in Supabase: Project Settings → API Keys).
2. Install dependencies: `npm install`
3. Run it locally: `npm run dev`

## Deploying

This project is built to deploy directly on [Vercel](https://vercel.com) — connect this GitHub repository to a new Vercel project, add the same two environment variables from `.env` into Vercel's Project Settings → Environment Variables, and deploy.

## Checking the database connection

Visit `your-deployed-url.vercel.app/?test` to see a simple connection test page, confirming the app can actually reach the Supabase database, before relying on the main app's login and data features.
