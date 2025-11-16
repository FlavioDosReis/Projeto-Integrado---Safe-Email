SafeMail IA – Analisador de E-mails

O SafeMail IA é um sistema desenvolvido em Python para análise automática de e-mails no Outlook Desktop, classificando mensagens em Baixo, Médio ou Alto Risco de acordo com heurísticas de segurança, palavras-chave, anexos suspeitos e URLs potencialmente maliciosas.

✅ Requisitos do Sistema

Antes de executar o projeto, certifique-se de que o ambiente possui:

Sistema Operacional
Windows 10 ou superior
(necessário para integração via MAPI/COM com o Outlook)

Softwares
Outlook Desktop instalado e configurado com uma conta ativa.
Python 3.9 ou superior

Bibliotecas Python necessárias
Instale usando:
pip install pywin32
(Outras bibliotecas como re, csv, difflib já fazem parte da biblioteca padrão do Python.)

📦 Instalação
1- Baixe ou clone o repositório do projeto:
git clone https://github.com/seu-repositorio/safemail-ia.git

2- Acesse a pasta do projeto:
cd safemail-ia

3- Instale a dependência principal:
pip install pywin32

▶️ Como Executar a Aplicação

1- Abra o terminal na pasta do projeto.
2- Execute o script principal:
python analisador_de_risco_outlook.py
3 -Certifique-se de que o Outlook esteja aberto ou configurado corretamente
(o script usa a interface MAPI via COM).

📊 Saída Gerada

Após a execução, o sistema irá:
Ler os e-mails da caixa Inbox (ou outra pasta configurada).
Calcular pontuação de risco.
Classificar cada e-mail em Baixo, Médio ou Alto risco.
Gerar um arquivo CSV contendo:
Data
Assunto
Remetente
Anexos
URLs
Palavras suspeitas
Pontuação
Classificação

O arquivo é salvo automaticamente na pasta do projeto.

⚙️ Configurações Ajustáveis

Dentro do código você pode configurar:
Pasta de e-mails a analisar (default: Inbox)
Número máximo de e-mails
Palavras-chave suspeitas
Extensões perigosas
Pesos das heurísticas
Se deseja marcar o assunto do e-mail com:
[Risco:ALTO] / [Risco:MÉDIO] / [Risco:BAIXO]

❗ Observações Importantes

O script não envia e-mails, apenas lê e marca mensagens.
Não depende de consultas externas (WHOIS, APIs, etc.).
Não modifica anexos, apenas os classifica.
A classificação é baseada em heurísticas simples e pode ser aprimorada com IA na próxima versão.

👨‍💻 Autores

Projeto desenvolvido pelos alunos:
Maicon Bruno Corrêa da Silva
Antonio Tiago Zaneratto
Flavio Perussi Bertão dos Reis
João Pedro Dutra da Silva
Gabriel Trinca de Marchi
