# 📦 Automação de Baixa de AFs via E-mail com Python + UiPath

Este projeto automatiza a rotina de leitura de e-mails com pedidos de baixa de *AFs (Autorização de Fornecimento)* e executa automaticamente um processo no *ERP Teknisa* através do *UiPath*.

## 🚀 Funcionalidades

- 📥 Monitoramento em tempo real de e-mails recebidos no Outlook.
- 🔍 Extração automática do número do restaurante e das AFs listadas no corpo do e-mail.
- 📊 Geração de uma planilha dados_af.xlsx com as AFs para o UiPath processar.
- 🖱 Execução automática do processo de baixa no ERP Teknisa via .bat.
- 📩 O e-mail original é:
  - Marcado como *lido*.
  - Respondido automaticamente com a mensagem: Baixa Ok!.

---

## 🧰 Tecnologias Utilizadas

- *Python 3.10*
- *UiPath Studio*
- *Outlook 2013 (32 bits)*
- *OpenPyXL (para salvar Excel)*
- *win32com (para integração com o Outlook)*
- *Regex (para tratar texto do e-mail)*

---

## 🗂 Estrutura do Projeto

📁 Testes
├── monitorar_emails_af.py # Script Python principal
├── executar_uipath.bat # Aciona o processo no UiPath
├── dados_af.xlsx # Planilha gerada automaticamente
└── projeto_uipath/ # Pasta com o projeto UiPath (.xaml, project.json)

yaml
Copiar
Editar

---

## ▶ Como usar

1. *Instale as dependências no Python*:

```bash
pip install pandas openpyxl pywin32
Configure o processo no UiPath para ler o dados_af.xlsx e fazer a baixa no ERP.

Configure o .bat (executar_uipath.bat) para chamar seu processo no UiPath Assistant:

bat
Copiar
Editar
"C:\Users\SeuUsuario\AppData\Local\Programs\UiPath\Studio\UiRobot.exe" run --file "C:\Caminho\Para\Seu\Projeto\Main.xaml"
Execute o script principal:

bash
Copiar
Editar
python monitorar_emails_af.py
🧠 Exemplo de Funcionamento
E-mail recebido:
yaml
Copiar
Editar
Assunto: Baixa de AF

Restaurante: 0003
AFs:
115380
115379

Saída no terminal:
yaml
Copiar
Editar
📧 Novo e-mail: Baixa de AF
📍 Restaurante: 0003
🧾 AFs: ['115380', '115379']
✅ Dados salvos no dados_af.xlsx
📩 E-mail marcado como lido.
✉ E-mail de confirmação enviado.
🚀 UiPath iniciado automaticamente!
🔐 Observações
O Outlook precisa estar aberto e configurado corretamente.

O processo no UiPath precisa estar testado e funcional.

Você pode adaptar o script para diferentes formatos de e-mail ou ERPs.

👨‍💻 Autor
Patrick – Analista e desenvolvedor de sistemas
📧 patrickbrando18102003@gmail.com
📍 Brasi