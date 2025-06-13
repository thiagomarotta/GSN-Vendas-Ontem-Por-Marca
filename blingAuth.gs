const BLING_CONFIG = {
    gsn: {
        clientId: "d6831c54ce7c590a0405df8f27669f1c7cb7a883",
        clientSecret: "389efc52b8a8659e431340289ec010c71e710504a582adf3c9995cb407df",
        tokenUrl: "https://www.bling.com.br/Api/v3/oauth/token",
        bling_redirect_uri: 'http://localhost:3000/callback'
    },
    metabolik: {
        clientId: "b72925ee320248cef9201ab5177da0f00f045dca",
        clientSecret: "d3255f85349806f0b64682a05d484c926e408ff18ac8cf051c91093e61c8",
        tokenUrl: "https://www.bling.com.br/Api/v3/oauth/token",
        bling_redirect_uri: 'http://localhost:3000/callback'
    }
};

function openAuthBlingAllAccounts() {
    const gerarLink = (prefix) => {
        const config = BLING_CONFIG[prefix];
        const state = 'bling_auth_' + new Date().getTime();
        PropertiesService.getScriptProperties().setProperty(`state_${prefix}`, state);

        return `https://www.bling.com.br/Api/v3/oauth/authorize` +
            `?response_type=code` +
            `&client_id=${config.clientId}` +
            `&redirect_uri=${encodeURIComponent(config.bling_redirect_uri)}` +
            `&state=${state}`;
    };

    const urlGSN = gerarLink('gsn');
    const urlMetabolik = gerarLink('metabolik');

    const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial, sans-serif; padding:10px;">
      <h2>🔗 Autenticação Bling - GSN e Metabolik</h2>

      <p style="color:darkred;"><b>⚠️ Importante:</b> Após autorizar uma conta, você deve <b>fazer logout no Bling</b> antes de autorizar a segunda conta.</p>

      <hr>
      <h3>🟦 Conta GSN</h3>
      <p><b>1️⃣ Clique no link abaixo para autorizar GSN:</b></p>
      <p><a href="${urlGSN}" target="_blank">➡️ <b>Autorizar Bling (GSN)</b></a></p>
      <p>Após autorizar, copie o <b>code</b> da URL e cole abaixo:</p>
      <input type="text" id="codigo_gsn" placeholder="Cole aqui o código da GSN" style="width:100%; padding:5px;">
      <br><br>
      <button onclick="enviar('gsn')">✅ Confirmar GSN</button>
      <p id="status_gsn" style="color:green; margin-top:5px;"></p>

      <hr>
      <h3>🟧 Conta Metabolik</h3>
      <p><b>1️⃣ Clique no link abaixo para autorizar Metabolik:</b></p>
      <p><a href="${urlMetabolik}" target="_blank">➡️ <b>Autorizar Bling (Metabolik)</b></a></p>
      <p>Após autorizar, copie o <b>code</b> da URL e cole abaixo:</p>
      <input type="text" id="codigo_metabolik" placeholder="Cole aqui o código da Metabolik" style="width:100%; padding:5px;">
      <br><br>
      <button onclick="enviar('metabolik')">✅ Confirmar Metabolik</button>
      <p id="status_metabolik" style="color:green; margin-top:5px;"></p>

      <hr>
      <p style="color:gray;">🚀 Após confirmar ambos, os tokens serão armazenados corretamente.</p>
    </div>

    <script>
      function enviar(prefix) {
        const input = document.getElementById('codigo_' + prefix);
        const status = document.getElementById('status_' + prefix);
        const codigo = input.value.trim();

        if (!codigo) {
          alert('Por favor, insira o código para ' + prefix.toUpperCase());
          return;
        }

        status.style.color = 'blue';
        status.innerText = 'Enviando...';

        google.script.run
          .withSuccessHandler(function(msg) {
            status.style.color = 'green';
            status.innerText = msg;
            alert(msg);
          })
          .withFailureHandler(function(error) {
            status.style.color = 'red';
            status.innerText = '❌ Erro: ' + error.message;
            alert('Erro: ' + error.message);
          })
          .processAndStoreCode(prefix, codigo);
      }
    </script>
  `).setWidth(600).setHeight(600);

    SpreadsheetApp.getUi().showModalDialog(html, `Autenticação Bling - GSN e Metabolik`);
}

function processAndStoreCode(prefix, codigo) {
    Logger.log(`🔁 processAndStoreCode chamado para ${prefix} com code: ${codigo}`);
    const config = BLING_CONFIG[prefix];
    if (!config) {
        throw new Error(`Prefixo "${prefix}" não encontrado na configuração.`);
    }

    const clientId = config.clientId;
    const clientSecret = config.clientSecret;

    const basicAuth = Utilities.base64Encode(`${clientId}:${clientSecret}`);

    const options = {
        method: "post",
        headers: {
            Authorization: `Basic ${basicAuth}`,
            "Content-Type": "application/x-www-form-urlencoded"
        },
        payload: {
            grant_type: "authorization_code",
            code: codigo,
            redirect_uri: config.bling_redirect_uri
        },
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(config.tokenUrl, options);
        const status = response.getResponseCode();
        const texto = response.getContentText();

        Logger.log(`Status: ${status}`);
        Logger.log(`Resposta: ${texto}`);

        const data = JSON.parse(texto);

        if (status !== 200) {
            throw new Error(`Erro na API: ${texto}`);
        }

        if (data.access_token) {

            Logger.log(`🔐 access_token recebido: ${data.access_token}`);
            Logger.log(`🔐 refresh_token recebido: ${data.refresh_token}`);
            Logger.log(`🔐 expires_in: ${data.expires_in}`);

            const props = PropertiesService.getScriptProperties();
            props.setProperty(`access_token_${prefix}`, data.access_token);
            props.setProperty(`refresh_token_${prefix}`, data.refresh_token);
            props.setProperty(`timestamp_${prefix}`, Date.now().toString());
            props.setProperty(`expires_in_${prefix}`, data.expires_in.toString());

            return '✅ Tokens obtidos e salvos com sucesso para ' + prefix.toUpperCase() + '!';
        } else {
            throw new Error('❌ Erro ao obter tokens: ' + texto);
        }
    } catch (e) {
        Logger.log('Erro: ' + e.message);
        throw new Error('❌ Erro na solicitação: ' + e.message);
    }
}

function getSavedBlingTokens(prefix) {
    const props = PropertiesService.getScriptProperties();
    return {
        access_token: props.getProperty(`access_token_${prefix}`),
        refresh_token: props.getProperty(`refresh_token_${prefix}`),
        timestamp: Number(props.getProperty(`timestamp_${prefix}`)),
        expires_in: Number(props.getProperty(`expires_in_${prefix}`)) || 21600
    };
}

function saveBlingTokens(prefix, tokens) {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`access_token_${prefix}`, tokens.access_token);
    props.setProperty(`refresh_token_${prefix}`, tokens.refresh_token);
    props.setProperty(`timestamp_${prefix}`, Date.now().toString());
    props.setProperty(`expires_in_${prefix}`, tokens.expires_in.toString());
}

function ensureValidBlingToken({ prefix, clientId, clientSecret, tokenUrl }) {
    const tokens = getSavedBlingTokens(prefix);
    const now = Date.now();
    const expiry = tokens.timestamp + tokens.expires_in * 1000;

    if (now < expiry) return tokens.access_token;

    const basicAuth = Utilities.base64Encode(`${clientId}:${clientSecret}`);
    const options = {
        method: "post",
        headers: {
            Authorization: `Basic ${basicAuth}`,
            "Content-Type": "application/x-www-form-urlencoded"
        },
        payload: `grant_type=refresh_token&refresh_token=${tokens.refresh_token}`,
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(tokenUrl, options);
    const data = JSON.parse(response.getContentText());

    if (data.access_token) {
        saveBlingTokens(prefix, data);
        return data.access_token;
    } else {
        Logger.log(response.getContentText());
        throw new Error(`Erro ao renovar o token do Bling (${prefix})`);
    }
}

function logSavedBlingTokens(prefix = "gsn") {
    const props = PropertiesService.getScriptProperties();

    const keys = [
        `access_token_${prefix}`,
        `refresh_token_${prefix}`,
        `timestamp_${prefix}`,
        `expires_in_${prefix}`
    ];

    keys.forEach(key => {
        const value = props.getProperty(key);
        Logger.log(`${key}: ${value}`);
    });
}

function getCompany(company = "gsn") {

    const token = company === "gsn" ? ensureValidBlingToken({ prefix: "gsn", ...BLING_CONFIG.gsn }) : ensureValidBlingToken({ prefix: "metabolik", ...BLING_CONFIG.metabolik });

    const url = "https://api.bling.com.br/Api/v3/empresas/me/dados-basicos";


    const options = {
        method: "get",
        headers: {
            "Authorization": `Bearer ${token}`,
            "Accept": "application/json"
        },
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());

        const nome = json?.data?.nome;
        if (!nome) throw new Error(`Nome da empresa não encontrado na resposta: ${JSON.stringify(json)}`);

        return nome;
    } catch (e) {
        Logger.log(`Erro ao obter dados da empresa: ${e.message}`);
        throw e;
    }
}