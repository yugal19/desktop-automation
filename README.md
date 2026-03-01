# 🎙️ Desktop Automation Through Voice
AI-powered desktop automation system that executes system-level actions using natural voice commands.

This project enables real-time voice-controlled desktop interaction using speech recognition and manually defined intent handling logic.
All automation logic is implemented purely in Python.

The system executes OS-level commands directly without AI-based intent parsing.
It uses FastAPI and WebSockets for real-time communication, and includes a separate server.py file exposing an MCP endpoint specifically for Excel automation.

---

## 🚀 Overview

Desktop Automation Through Voice allows users to:

- 🎤 Speak commands naturally  
- 🧠 Convert speech → structured intent  
- ⚡ Execute desktop-level actions in real time  
- 🌐 Interact via web interface + WebSockets  
- 🖥️ Automate system workflows hands-free  

---

## 🧠 System Architecture

### Voice Flow

User Voice  
→ Speech-to-Text  
→ Interpreter (Intent Parsing)  
→ Action Mapper  
→ Desktop Automation (OS / subprocess / automation scripts)

### Web Flow

Browser (form.html)  
→ FastAPI Server (main.py / server.py)  
→ WebSocket Layer (web_socket_server.py)  
→ Interpreter  
→ actions.py  
→ System Execution  



## 📂 Project Structure

```
desktop-automation-voice/
├── .gitignore
├── actions.py 
├── claude_writer.py
├── form.html
├── interpreter.py
├── main.py
├── server.py
├── web_form_controller.py
├── web_socket_server.py
├── requirements.txt
```


## Installation And How to Execute

### 1. Clone the Repository
```bash
git clone <repo-link>
cd <repo-folder>
```

### 2. Create Virtual Environment
```bash
python -m venv venv
```

### 3. Activate Virtual Environment

Windows:
```bash
venv\Scripts\activate
```

Mac / Linux:
```bash
source venv/bin/activate
```

### 4. Install Requirements
```bash
pip install -r requirements.txt
```

### 5. Run the Application
```bash
python main.py
```


## Demo of the Project :- 

https://github.com/user-attachments/assets/5931769b-9ced-4c58-aa4f-1b1cba7efac5



