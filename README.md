
# CCR Seguimiento – Render Deployment

## 1. Required Environment Variables (Render)
- ITEM_ID = ID of your Excel file in OneDrive

## 2. Setup in Azure App Registration
- Type: Accounts in personal Microsoft accounts
- Redirect URI:
  https://formulario-ccr-jjrg-160126-v1.onrender.com/auth/callback
- API Permissions:
  - Files.ReadWrite
  - User.Read
  - offline_access
- Create a Client Secret
  
## 3. Upload to GitHub
Include:
- app.py  
- formulario.html  
- config.json  
- requirements.txt  
- Procfile  

## 4. Deploy on Render
- Build command: pip install -r requirements.txt
- Start command: gunicorn app:app

