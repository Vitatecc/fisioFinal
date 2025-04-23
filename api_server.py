from fastapi import FastAPI
import subprocess

app = FastAPI()

@app.get("/")
def home():
    return {"message": "FisioAutomatizacion API online 🚀 (v2)"}

@app.get("/extraer-citas")
def extraer_citas():
    result = subprocess.run(["python3", "extraer_citas.py"], capture_output=True, text=True)
    return {
        "stdout": result.stdout,
        "stderr": result.stderr
    }

@app.get("/crear-usuario")
def crear_usuario():
    result = subprocess.run(["python3", "Crear_usuario.py"], capture_output=True, text=True)
    return {
        "stdout": result.stdout,
        "stderr": result.stderr
    }
