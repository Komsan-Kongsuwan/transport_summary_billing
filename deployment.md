# Quick Deployment Guide

## Option 1: Streamlit Cloud (Recommended - FREE)

### Step 1: Prepare Your Files
You need these 3 files:
- `billing_summary_app.py` (main app)
- `requirements.txt` (dependencies)
- `.streamlit/config.toml` (optional, for better UI)

### Step 2: Create GitHub Repository

1. Go to https://github.com/new
2. Create a new repository (e.g., "billing-summary-app")
3. Upload the files:
   - Click "uploading an existing file"
   - Drag and drop all three files
   - Commit changes

### Step 3: Deploy to Streamlit Cloud

1. Go to https://share.streamlit.io/
2. Click "Sign in with GitHub"
3. Authorize Streamlit
4. Click "New app" button
5. Fill in the form:
   - **Repository**: Select your repository
   - **Branch**: main
   - **Main file path**: `billing_summary_app.py`
6. Click "Deploy!"
7. Wait 2-3 minutes for deployment

### Step 4: Get Your App URL
- You'll receive a URL like: `https://billing-summary-app.streamlit.app`
- Share this URL with your team!

### Step 5: Update Your App
To update the app:
1. Edit files on GitHub
2. Commit changes
3. Streamlit Cloud auto-deploys (no action needed!)

---

## Option 2: Local Deployment

### For Windows:

1. **Install Python** (if not installed)
   - Download from https://www.python.org/downloads/
   - Check "Add Python to PATH" during installation

2. **Open Command Prompt**
   - Press Win+R, type `cmd`, press Enter

3. **Navigate to app folder**
   ```cmd
   cd path\to\your\app\folder
   ```

4. **Install dependencies**
   ```cmd
   pip install -r requirements.txt
   ```

5. **Run the app**
   ```cmd
   streamlit run billing_summary_app.py
   ```

6. **Access the app**
   - Browser will open automatically at http://localhost:8501
   - If not, open browser and go to http://localhost:8501

### For Mac/Linux:

1. **Open Terminal**

2. **Navigate to app folder**
   ```bash
   cd /path/to/your/app/folder
   ```

3. **Install dependencies**
   ```bash
   pip3 install -r requirements.txt
   ```

4. **Run the app**
   ```bash
   streamlit run billing_summary_app.py
   ```

5. **Access the app**
   - Browser will open automatically at http://localhost:8501

---

## Option 3: Heroku (Alternative Cloud)

### Prerequisites:
- Heroku account (free): https://heroku.com
- Git installed

### Steps:

1. **Create Procfile**
   Create a file named `Procfile` (no extension) with:
   ```
   web: streamlit run billing_summary_app.py --server.port=$PORT --server.address=0.0.0.0
   ```

2. **Create setup.sh**
   Create a file named `setup.sh` with:
   ```bash
   mkdir -p ~/.streamlit/
   echo "\
   [server]\n\
   headless = true\n\
   port = $PORT\n\
   enableCORS = false\n\
   \n\
   " > ~/.streamlit/config.toml
   ```

3. **Update Procfile to use setup.sh**
   ```
   web: sh setup.sh && streamlit run billing_summary_app.py
   ```

4. **Deploy**
   ```bash
   # Login to Heroku
   heroku login
   
   # Create app
   heroku create your-app-name
   
   # Initialize git (if not done)
   git init
   git add .
   git commit -m "Initial commit"
   
   # Deploy
   git push heroku main
   
   # Open app
   heroku open
   ```

---

## Troubleshooting

### Streamlit Cloud Issues

**Problem: App doesn't start**
- Solution: Check requirements.txt has correct package names
- Check logs in Streamlit Cloud dashboard

**Problem: File upload fails**
- Solution: File might be too large (>200MB limit)
- Try with smaller file or use local deployment

**Problem: App is slow**
- Solution: Free tier has limited resources
- Consider upgrading to paid tier or use local deployment

### Local Issues

**Problem: "streamlit: command not found"**
- Solution (Windows): 
  ```cmd
  python -m streamlit run billing_summary_app.py
  ```
- Solution (Mac/Linux):
  ```bash
  python3 -m streamlit run billing_summary_app.py
  ```

**Problem: "Module not found"**
- Solution: Make sure you installed requirements:
  ```bash
  pip install -r requirements.txt
  ```

**Problem: Port already in use**
- Solution: Kill existing Streamlit process or use different port:
  ```bash
  streamlit run billing_summary_app.py --server.port=8502
  ```

---

## Network Access

### For Local Deployment on Company Network:

1. **Find your local IP**
   - Windows: `ipconfig` (look for IPv4 Address)
   - Mac/Linux: `ifconfig` (look for inet)

2. **Share with colleagues**
   - Example: `http://192.168.1.100:8501`
   - Colleagues on same network can access

3. **Firewall Settings**
   - You may need to allow port 8501 through firewall
   - Windows: Windows Defender Firewall > Advanced Settings > Inbound Rules
   - Mac: System Preferences > Security & Privacy > Firewall

---

## Best Practices

### Security
- Don't commit sensitive data to GitHub
- Use environment variables for sensitive configs
- Streamlit Cloud apps are public by default

### Performance
- Test with small files first
- Consider file size limits (200MB for Streamlit Cloud)
- Use local deployment for very large files

### Maintenance
- Keep dependencies updated
- Monitor app performance
- Check Streamlit Cloud logs regularly

---

## Getting Help

- **Streamlit Docs**: https://docs.streamlit.io
- **Streamlit Community**: https://discuss.streamlit.io
- **GitHub Issues**: Create issue in your repository

---

## Quick Reference

### Useful Commands

```bash
# Run app locally
streamlit run billing_summary_app.py

# Run on specific port
streamlit run billing_summary_app.py --server.port=8502

# Open app in browser
# (automatically opens, or go to http://localhost:8501)

# Stop app
# Press Ctrl+C in terminal

# Update dependencies
pip install -r requirements.txt --upgrade

# Clear Streamlit cache
# Click "Clear cache" in app menu (top right)
```

### File Structure
```
billing-summary-app/
├── billing_summary_app.py    # Main app
├── requirements.txt           # Dependencies
├── .streamlit/
│   └── config.toml           # Config (optional)
├── README.md                 # Documentation
└── DEPLOYMENT.md            # This file
```

---

## Success! 🎉

Once deployed, your app will:
- ✅ Be accessible 24/7
- ✅ Auto-update when you push changes
- ✅ Handle multiple users
- ✅ Process billing data automatically
- ✅ Generate downloadable Excel reports

Share your app URL with your team and start processing billing data in the cloud!
