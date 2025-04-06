from flask import Flask , render_template_string, request, redirect, jsonify
import json
import os
import pyautogui
import webbrowser
import pythoncom
import comtypes
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

app = Flask(__name__)
ARQUIVO_JSON = "atalhos.json"


# Cria o JSON se não existir
if not os.path.exists(ARQUIVO_JSON):
    with open(ARQUIVO_JSON, "w") as f:
        json.dump([], f)

# Função para carregar atalhos do JSON
def carregar_atalhos():
    with open(ARQUIVO_JSON, "r") as f:
        return json.load(f)

# Função para salvar atalhos no JSON
def salvar_atalhos(lista):
    with open(ARQUIVO_JSON, "w") as f:
        json.dump(lista, f, indent=4)


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nome = request.form.get("name")
        url = request.form.get("url")
        local = request.form.get("localFile")

        atalhos = carregar_atalhos()
        atalhos.append({"nome": nome, "url": url, "local": local})
        salvar_atalhos(atalhos)

        return redirect("/")

    atalhos = carregar_atalhos()
    return render_template_string(HTML_TEMPLATE, atalhos=atalhos)

@app.route("/excluir/<nome>", methods=["POST"])
def excluir(nome):
    atalhos = carregar_atalhos()
    atalhos = [a for a in atalhos if a["nome"] != nome]
    salvar_atalhos(atalhos)
    return redirect("/")


# HTML será colocado aqui abaixo ↓
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Hub Control Server</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
</head>
<body class="bg-dark text-light">
<div class="container-fluid px-3">
    <h1 class="text-center mb-4 fs-3">Hub de Atalhos</h1><br>
    <h3 class="text-center mb-4 fs-3">Adiconar Novo Atalho</h3><br>
    <form class="row g-4" action="/" method="post">
        <div class="col-12 col-md-6">

            <label for="name" class="form-label">Nome do atalho</label>
            <input type="text" class="form-control" name="name" required>
        </div>
        <div class="col-12 col-md-6">

            <label for="url" class="form-label">Url do atalho</label>
            <input type="text" class="form-control" name="url">
        </div>
        <div class="col-12 col-md-6">

            <label for="localFile" class="form-label">Arquivo local</label>
            <input type="text" class="form-control" name="localFile">
        </div>
        <div class="col-md-12">
            <button type="submit" class="btn btn-primary mt-3">Salvar</button>
        </div>
    </form><br>
    <h3 class="text-center mb-4 fs-3">Atalhos Cadastrados</h3><br>
    <p class="text-center">Clique em "Executar" para abrir o atalho ou "Excluir" para removê-lo.</p>
    <ul class="list-group">
        {% for atalho in atalhos %}
        <li class="list-group-item bg-dark text-white d-flex flex-column flex-md-row justify-content-between align-items-start align-items-md-center gap-2">
            <div>
                <strong>{{ atalho.nome }}</strong><br>
                {% if atalho.url %}
                    <a href="{{ atalho.url }}" class="text-info" target="_blank">{{ atalho.url }}</a><br>
                {% endif %}
                {% if atalho.local %}
                    <span class="text-white">{{ atalho.local }}</span>
                {% endif %}
            </div>
            <div class="d-flex gap-2">
                <button class="btn btn-success btn-sm" onclick="executarAtalho('{{ atalho.nome }}')">Executar</button>
                <script>
                    function executarAtalho(nome) 
                    {
                        fetch(`/open/${nome}`)
                            .then(response => response.text())
                            .then(data => {
                                console.log("Executado:", data);
                                // opcional: mostrar um toast ou alerta
                                // alert("Atalho executado com sucesso!");
                            })
                            .catch(error => {
                                alert("Erro ao executar atalho: " + error);
                                // opcional: mostrar um toast ou alerta
                                console.error("Erro ao executar atalho:", error);
                            });
                    }                    
                </script>
                <form action="/excluir/{{ atalho.nome }}" method="post" style="margin: 0;">
                    <button class="btn btn-danger btn-sm">Excluir</button>
                </form>
            </div>
        </li>
        {% endfor %}
    </ul>
</div>
</body>
</html>
"""


# Função para obter o dispositivo de áudio padrão
def get_audio_volume():
    pythoncom.CoInitialize()  # Inicializa o COM
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
    return interface.QueryInterface(IAudioEndpointVolume)

@app.route("/open/<nome>", methods=["GET"])
def executar_atalho(nome):
    atalhos = carregar_atalhos()
    for atalho in atalhos:
        if atalho.get("nome") == nome:
            url = atalho.get("url", "")
            local = atalho.get("local", "")
            try:
                if url:
                    webbrowser.open(url)
                    return f"Abrindo URL: {url}"
                elif local:
                    os.startfile(local)
                    return f"Abrindo arquivo: {local}"
                else:
                    return "Nenhum caminho definido para este atalho."
            except Exception as e:
                return f"Erro ao executar o atalho: {str(e)}"
    return f"Atalho '{nome}' não encontrado."


# ------------------------
# Rotas fixas do seu código
# ------------------------

@app.route('/Vol+', methods=['GET'])
def increase_volume():
    volume = get_audio_volume()
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = min(current_volume + 0.05, 1.0)
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    return f'Volume aumentado para {new_volume * 100:.0f}%'

@app.route('/Vol-', methods=['GET'])
def decrease_volume():
    volume = get_audio_volume()
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = max(current_volume - 0.05, 0.0)
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    return f'Volume diminuído para {new_volume * 100:.0f}%'

@app.route('/mute', methods=['GET'])
def mute_volume():
    volume = get_audio_volume()
    volume.SetMute(True, None)
    return 'Volume mutado'

@app.route('/unmute', methods=['GET'])
def unmute_volume():
    volume = get_audio_volume()
    volume.SetMute(False, None)
    return 'Volume desmutado'

@app.route('/PwOff', methods=['GET'])
def shutdown():
    os.system('shutdown /s /t 1')
    return 'Shutting down...'

@app.route('/open/Config', methods=['GET'])
def open_settings():
    os.system('start ms-settings:')
    return 'Configurações abertas'

@app.route('/Lock', methods=['GET'])
def lock_pc():
    os.system('rundll32.exe user32.dll,LockWorkStation')
    return 'PC bloqueado'

@app.route('/music/Play-Pause', methods=['GET'])
def play_music():
    pyautogui.press('playpause')
    return 'Toggled play/pause'

@app.route('/music/Next', methods=['GET'])
def next_music():
    pyautogui.press('nexttrack')
    return 'Skipped to next track'

@app.route('/music/Prev', methods=['GET'])
def previous_music():
    pyautogui.press('prevtrack')
    return 'Went to previous track'


# Inicia o servidor Flask
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5100)
