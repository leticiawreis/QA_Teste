# QA Control Center

Aplicação desktop em Python + PyQt6 para organizar atividades de QA de RPAs.

## Instalação

```
python -m pip install -r requirements.txt
python main.py
```

## Onde os dados ficam

O sistema cria automaticamente, na mesma pasta onde está o `main.py` (ou o
executável `.exe`, se você estiver usando a versão empacotada), a pasta:

```
QA/
```

com as subpastas:

* 01 - Dashboard
* 02 - Casos de Teste
* 03 - Bugs
* 04 - Evidências
* 05 - Base de Conhecimento
* 06 - Relatórios Mensais
* 07 - Riscos
* 08 - Checklists

Como a pasta é criada ao lado do programa, se você mover o projeto (ou o
`.exe`) para outro lugar, os dados vão junto.

## Armazenamento

Não usa banco de dados. Cada categoria (Casos de Teste, Bugs, Riscos, Base de
Conhecimento, Checklists) tem **um único arquivo Excel** dentro da sua
respectiva pasta, e cada registro feito pela interface vira uma nova linha
nesse Excel, automaticamente:

* `02 - Casos de Teste/Casos de Teste.xlsx`
* `03 - Bugs/Bugs.xlsx`
* `07 - Riscos/Riscos.xlsx`
* `05 - Base de Conhecimento/Base de Conhecimento.xlsx`
* `08 - Checklists/Checklists.xlsx`

Evidências anexadas (prints, arquivos, etc.) são copiadas para
`04 - Evidências/<ID-do-registro>/`.

## Fluxo

1. Abra o programa.
2. Escolha o tipo de registro no menu lateral.
3. Preencha os campos.
4. Anexe evidência quando necessário.
5. Clique em Registrar.
6. O ID é criado automaticamente e a linha é adicionada ao Excel da categoria.
7. O Dashboard é atualizado com os novos números.

## Gerando o executável (.exe)

```
python build.py
```

Isso instala o PyInstaller (se necessário) e gera o executável em
`dist/QA Control Center.exe`. Depois é só mover esse `.exe` para a pasta onde
você quer que o app (e a pasta `QA` com os dados) fique.

## Próximas evoluções recomendadas

* edição/exclusão de registros;
* pesquisa e filtros;
* exportação para PDF;
* gráficos mensais;
* cadastro de usuários/RPAs;
* tela de relatórios consolidados;
* controle de permissões;
* backup automático das planilhas.