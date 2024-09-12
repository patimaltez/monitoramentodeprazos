Documentação do Projeto: Gerenciamento e Alerta de Prazo de Correção de Provas
Descrição
Este projeto consiste em um script Python desenvolvido para gerenciar arquivos PDF e enviar alertas ao professor sobre prazos de correção. O código faz o seguinte:

Copia arquivos PDF das pastas de origem (Desktop/TCC e Downloads/provas bimestrais) para pastas de destino específicas (Documents/curso ciencia da computação/TCC e Documents/curso ciencia da computação/provas bimestrais).
Remove arquivos PDF das pastas de origem após a cópia.
Envia alertas por e-mail através do Microsoft Outlook, notificando sobre o início do prazo de correção, a metade do prazo e o alerta final quando faltar 24 horas para o prazo ser concluído.
Requisitos
Python:

Python 3.12 ou superior.
Bibliotecas Python necessárias:
pywin32 para integração com o Outlook.
os, shutil, datetime para manipulação de arquivos e datas.
Microsoft Outlook:

O Microsoft Outlook deve estar instalado e configurado corretamente no computador onde o script será executado.
Permissões de automação devem estar habilitadas no Outlook.
Configurações do Sistema:

Diretórios:
Diretório de origem para TCC: C:\Users\Patricia\Desktop\TCC
Diretório de origem para provas bimestrais: C:\Users\Patricia\Downloads\provas bimestrais
Diretórios de destino: C:\Users\Patricia\Documents\curso ciencia da computação\TCC e C:\Users\Patricia\Documents\curso ciencia da computação\provas bimestrais
Permissões e Configurações de Segurança:

Verifique se o Outlook permite acesso via automação. Ajuste as permissões de segurança se necessário.
O script requer acesso de leitura e escrita aos diretórios especificados.
Como Executar
Preparação:

Certifique-se de que todos os diretórios e arquivos necessários estão configurados corretamente.
Instale as dependências necessárias com pip install pywin32.
Execução:

Execute o script diretamente no ambiente Python ou configure um arquivo .bat para execução automática.
Configuração do Trigger:

Embora o agendamento de tarefas não tenha sido configurado, a execução manual do script garantirá que os arquivos sejam gerenciados e os alertas sejam enviados conforme esperado.
