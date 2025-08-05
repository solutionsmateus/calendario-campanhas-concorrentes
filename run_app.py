import pandas
import numpy
from flask import Flask


SCRIPT_RUN_DIR = "/app.py"

app = Flask(__name__)
    
def run_app():
    methods = [getattr[filter(SCRIPT_RUN_DIR)]]
    for method in methods():
        if method == []:
            return "Metodos não encontrados"
        
    

if __name__ == "__main__":
    app.run(debug=True)

