# Repositório de Código IPYNB para Mover Arquivos e Emitir Alertas de Prazos via Outlook

Este repositório contém um notebook Jupyter (.ipynb) que realiza as seguintes tarefas:

1. **Movimentação de Arquivos PDF**: Move arquivos PDF de diretórios específicos para uma pasta de destino, organizando-os conforme a categoria (TCC e provas bimestrais). Além disso, evita a duplicação de arquivos já existentes na pasta de destino e exclui os arquivos do diretório de origem após a cópia.

2. **Emissão de Alertas de Prazos**: Envia alertas via e-mail usando o Outlook para notificar o interessado sobre o início do prazo de correção, a metade do prazo e a proximidade do prazo final para a entrega das notas.

## Requisitos

Para garantir que o código funcione corretamente, você precisa atender aos seguintes requisitos:

1. **Python**: Certifique-se de ter o Python 3.12 ou superior instalado em seu computador.

2. **Pacotes Python**: Instale as seguintes bibliotecas Python:
   - `pywin32` para integração com o Outlook.
   - `os`, `shutil`, `datetime` para manipulação de arquivos e datas.

3. **Microsoft Outlook**: Instale as seguintes bibliotecas Python:
   - O Microsoft Outlook deve estar instalado e configurado corretamente no computador onde o script será executado.
   - Permissões de automação devem estar habilitadas no Outlook.

4. **Configurações do Sistema**: Diretórios de acordo com nosso exemplo
   - Diretório de origem para TCC: C:\Users\Name\Desktop
   - Diretório de origem para provas bimestrais: C:\Users\Name\Downloads
  
5. **Permissões e Configurações de Segurança**:
   - Verifique se o Outlook permite acesso via automação. Ajuste as permissões de segurança se necessário.
   - O script requer acesso de leitura e escrita aos diretórios especificados.

## Como Executar

Preparação:
 - Certifique-se de que todos os diretórios e arquivos necessários estão configurados corretamente.
 - Instale as dependências necessárias com pip install pywin32.

Execução:
 - Execute o script diretamente no ambiente Python ou configure um arquivo .bat para execução automática.

Configuração do Trigger:
 - O agendamento de tarefas não foi configurado, portanto execute manualmente o script.
 - Para criar um Trigger, veja o repositório patimaltez/criandotriggerparaipynb

## Conclusão

Com este projeto, você pode gerenciar eficientemente arquivos e garantir que prazos de retorno sobre eles sejam cumpridos.
Siga as instruções acima para garantir que o código funcione conforme esperado.
