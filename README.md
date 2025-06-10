# PPE Violation Detection Web App

This is a simple Flask-based web application to demonstrate PPE violation detection for helmets. It provides basic user management, image upload, IP camera capture, and records violations when helmets are not detected.

## Features

- User registration and login with roles
- Upload images or capture from IP cameras
- Placeholder helmet detection logic
- Violation archive with timestamps and associated users

## Running

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Start the application:
   ```bash
   python app.py
   ```
3. Access `http://localhost:5000` in your browser.

This project uses a placeholder detection function. Integrate an actual helmet detection model in `ppe_app/utils.py` by replacing the `detect_helmet` function.
