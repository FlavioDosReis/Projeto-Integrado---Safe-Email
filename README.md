🛡️ SafeMail IA – Analisador de E-mails
🔍 Análise Inteligente de Risco em Mensagens do Outlook



📘 Sobre o Projeto

O SafeMail IA é um analisador automático de risco para e-mails do Outlook Desktop, desenvolvido em Python.
Ele identifica mensagens suspeitas, analisa anexos, URLs, padrões de phishing e aplica uma classificação baseada em risco (Baixo, Médio, Alto).

O objetivo é aumentar a segurança corporativa, reduzir exposição a golpes e auxiliar usuários na tomada de decisão antes de abrir e-mails suspeitos.

Vídeo : https://www.youtube.com/watch?v=GhhBC6kXxUc

🚀 Funcionalidades

Leitura automática de e-mails via Outlook (MAPI/COM)

Detecção de padrões suspeitos:

palavras-chave maliciosas

URLs duvidosas

anexos perigosos

discrepâncias de remetente

Pontuação heurística de risco (0 a 100)

Classificação automática:

Baixo risco

Médio risco

Alto risco

Geração de relatório CSV detalhado

Marcação automática no assunto do e-mail (opcional)



🧩 Tecnologias Utilizadas

Tecnologia	              -    Finalidade

Python 3.9+	              -   Desenvolvimento principal

PyWin32	                  -    Integração COM com Outlook

difflib	                  -    Detecção de similaridade

Regex (re)	              -    Análise de URLs e padrões

CSV	                      -    Exportação de relatórios

Outlook Desktop	          -    Origem dos e-mails analisados



🔧 Requisitos

Sistema
Windows 10/11
Outlook Desktop configurado
Python 3.9+ instalado

Instalação de dependências
pip install pywin32



📦 Instalação do Projeto

Clone o repositório:

git clone https://github.com/seu-repositorio/safemail-ia.git


Acesse a pasta:

cd safemail-ia

Instale as dependências:

pip install pywin32



▶️ Como Executar

Execute o script principal:

python analisador_de_risco_outlook.py

Certifique-se de que o Outlook esteja aberto ou configurado no Windows, pois o script acessa a caixa de entrada via MAPI.




📊 Saídas do Sistema

O script gera:

✔ Relatório resultados.csv contendo:

data

remetente

assunto

anexos

URLs

palavras suspeitas

pontuação

classificação final

✔ Marcação no assunto:
[Risco:ALTO] Assunto original



⚙️ Configurações

Dentro do código, você pode ajustar:

Pasta alvo do Outlook

Número máximo de e-mails

Pesos das heurísticas

Lista de palavras suspeitas

Extensões perigosas

Ativar/desativar marcação no assunto



🧪 Testes Realizados

Outlook Desktop com conta ativa

Teste com e-mails reais e simulados

Links falsos (texto vs. URL real)

Anexos perigosos (.exe, .js, .docm, etc.)

E-mails corporativos legítimos

Performance com +500 mensagens



👨‍💻 Autores

Equipe de desenvolvimento:

Maicon Bruno Corrêa da Silva R.A: 24000795

Antonio Tiago Zaneratto R.A: 24000696

Flavio Perussi Bertão dos Reis Reis RA: 24001465

João Pedro Dutra da Silva RA: 24000990

Gabriel Trinca de Marchi RA: 24002112



📈 Melhorias Futuras

🤖 Implementação de rede neural ou modelo ML real

🖥 Interface gráfica (dashboard de risco)

📧 Compatibilidade com Gmail API

🔍 Análise profunda de anexos (sandboxing)

🧬 Algoritmos avançados de classificação
