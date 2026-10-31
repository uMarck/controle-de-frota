# Controle de Pranchas Alencar

Ferramenta estática para cruzar a planilha de controle, o mapa semanal de movimentações e o patiamento da empresa. A leitura e o cruzamento são executados no navegador do usuário. Arquivos selecionados não são enviados pelo aplicativo a um servidor.

## Arquivos do projeto

- `dist/`: página, scripts, leitor de planilhas e modelo vazio de exportação.
- `vercel.json`: informa à Vercel que o conteúdo publicado está em `dist`.
- `logic-check.cjs`: verificações da lógica de classificação.

## Publicar no GitHub e Vercel

1. Extraia este pacote e crie um repositório no GitHub.
2. Envie todos os arquivos e pastas extraídos para a raiz do repositório.
3. Na Vercel, escolha **Add New → Project**, importe esse repositório e mantenha as configurações de `vercel.json`.
4. Se a Vercel solicitar os campos manualmente, escolha **Other** como framework, deixe o comando de build vazio e defina `dist` como diretório de saída.
5. Clique em **Deploy**. Nos próximos envios ao GitHub, a Vercel publica as alterações automaticamente.

## Usar

Abra o endereço da Vercel, arraste as três planilhas, confira a aba correta de cada arquivo e clique em **Cruzar dados**. O Excel exportado segue o modelo de controle vazio e inclui o gráfico de pizza e as abas de conferência.

Use os arquivos da semana que pretende analisar. O patiamento usa a aba `PRANCHAS`; verde indica que aguenta peso, cinza que não aguenta e vermelho manutenção. Linhas vermelhas sem número de frota aparecem como pendências sem serem atribuídas a outra prancha.

## Privacidade dos dados

O pacote não contém as planilhas operacionais usadas como exemplo nem um resultado gerado a partir delas. Não envie as planilhas para um repositório público. Para um repositório privado, limite o acesso a pessoas autorizadas.

Os arquivos de planilha usados em cada sessão ficam no navegador enquanto a página está aberta. O site não tem autenticação, armazenamento de resultados, banco de dados ou servidor de processamento. O endereço da Vercel poderá ser acessado por qualquer pessoa que tenha o link, conforme as configurações de acesso da Vercel.
