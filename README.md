# AFSUR26 - Safety Updates

A professional Pharmacovigilance signal detection and ICSR overview tool.

## Local Development

To run this project locally on your Mac:

1. **Download the code**: Use the **Export to ZIP** option in the **Settings** menu of AI Studio.
2. **Unzip** the downloaded folder.
3. **Open Terminal** and navigate to the project folder.
4. **Install Dependencies**:
   ```bash
   npm install
   ```
5. **Run Development Server**:
   ```bash
   npm run dev
   ```
6. **Open your browser** at `http://localhost:3000`.

## Building for Production

To create a production-ready build:

```bash
npm run build
```

The output will be in the `dist/` directory.

## GitHub Pages Deployment

This project includes a GitHub Action in `.github/workflows/deploy.yml` that automatically deploys your app to GitHub Pages when you push to the `main` branch.
