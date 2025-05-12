const express = require('express');
const app = express();
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');

app.use(express.json());
app.use(express.static('public'));

const PORT = 3000;

// Objeto que armazena os status dos e-mails
let emailsStatus = {};

// Função para ler a planilha Excel
function lerPlanilhaExcel() {
    const workbook = xlsx.readFile('emails.xlsx');
    const planilha = workbook.Sheets[workbook.SheetNames[0]];
    const dados = xlsx.utils.sheet_to_json(planilha);
    return dados;
}

// Rota atualizada que junta planilha com os status atualizados
app.get('/contatos', async (req, res) => {
    try {
        const contatos = lerPlanilhaExcel();

        const contatosComStatus = contatos.map(c => {
            const email = (c.email || '').toLowerCase().trim();
            const statusInfo = emailsStatus[email];

            return {
                ...c,
                status: statusInfo ? statusInfo.status : 'não enviado'
            };
        });

        res.json(contatosComStatus);
    } catch (err) {
        console.error('Erro ao ler planilha:', err);
        res.status(500).json({ error: 'Erro ao ler planilha' });
    }
});

// Rota que recebe clique do pixel de rastreamento
app.get('/pixel.png', (req, res) => {
    const { email } = req.query;

    if (email) {
        const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress;
        const now = new Date().toISOString();

        const emailNormalizado = email.toLowerCase().trim();

        emailsStatus[emailNormalizado] = {
            status: 'visualizado',
            ip,
            data: now
        };

        console.log(`📬 E-mail aberto por: ${emailNormalizado} - IP: ${ip} - ${now}`);

        // Salvar os dados em arquivo opcionalmente:
        // fs.writeFileSync('status.json', JSON.stringify(emailsStatus, null, 2));
    }

    // Enviar imagem transparente (1x1)
    const imgPath = path.join(__dirname, 'public/pixel.png');
    res.sendFile(imgPath);
});

// Rota para atualizar status manualmente (ex: ao clicar "Finalizado")
app.post('/atualizar-status/:codEmpresa', (req, res) => {
    const { codEmpresa } = req.params;
    const { status } = req.body;

    // você pode atualizar `emailsStatus` aqui se quiser associar com o código da empresa
    res.sendStatus(200);
});

// Rota para resetar tudo
app.post('/resetar-tudo', (req, res) => {
    emailsStatus = {};
    res.sendStatus(200);
});

app.listen(PORT, () => {
    console.log(`Servidor rodando em http://localhost:${PORT}`);
});
