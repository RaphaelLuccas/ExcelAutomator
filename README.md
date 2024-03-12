Automação de Entrada de Dados de Produtos
Este script Python automatiza a entrada de dados de produtos em um sistema. Ele lê os dados de um arquivo Excel e usa a biblioteca pyautogui para interagir com a interface do usuário do sistema, inserindo os dados do produto nos campos apropriados.

Dependências
O script depende das seguintes bibliotecas Python:

openpyxl: para ler dados do arquivo Excel.
pyperclip: para copiar os dados para a área de transferência.
pyautogui: para automatizar a interação com a interface do usuário.
time: para introduzir atrasos entre as ações.
  
Como usar
Certifique-se de que todas as dependências estão instaladas.
Coloque o arquivo ‘produtos_ficticios.xlsx’ no mesmo diretório do script.
Execute o script.

Detalhes do Código
O script lê os dados do arquivo ‘produtos_ficticios.xlsx’ e itera sobre cada linha (produto). Para cada produto, ele copia cada campo (nome do produto, descrição, categoria, etc.) para a área de transferência e, em seguida, usa pyautogui para colar os dados no campo apropriado na interface do usuário.

Aviso
Este script foi projetado para funcionar com uma configuração específica de interface do usuário e pode não funcionar corretamente se a interface do usuário for alterada. Use com cuidado e sempre supervisione a execução do script para garantir que ele está funcionando como esperado.

Contribuições
Contribuições para este projeto são bem-vindas. Se você encontrar um bug ou tem uma sugestão de melhoria, sinta-se à vontade para abrir uma issue ou enviar um pull request.
