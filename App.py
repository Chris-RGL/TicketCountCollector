import webbrowser

from flask import Flask, render_template, request, jsonify, session
from threading import Thread
import time

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Required for session handling

# Global variable to store login information
login_info = []

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    username = request.form['username']
    password = request.form['password']
    global login_info
    login_info = [username, password]
    print(f"Username: {username}, Password: {password}")  # Print to console

    session.clear()
    return jsonify(status='success')

def run_flask():
    app.run(debug=False, use_reloader=False)  # Disable reloader and debug mode

def open_browser():
    time.sleep(1)  # Give Flask a moment to start up
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == "__main__":
    # Start the Flask app in a separate thread
    flask_thread = Thread(target=run_flask)
    flask_thread.start()
    # Open the web browser
    open_browser()