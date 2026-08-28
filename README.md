# QA Control Center

Aplicação desktop em Python + PyQt6 para organizar atividades de QA de RPAs.

## Instalação

```bash
python -m pip install -r requirements.txt
python main.py
```

O sistema cria automaticamente:

`~/QA_Control_Center/QA/`

com as pastas:
- 01 - Dashboard
- 02 - Casos de Teste
- 03 - Bugs
- 04 - Evidências
- 05 - Base de Conhecimento
- 06 - Relatórios Mensais
- 07 - Riscos
- 08 - Checklists

Os registros ficam em um banco SQLite local (`qa.db`) e também são exportados como arquivos `.txt` nas pastas correspondentes.

## Fluxo

1. Abra o programa.
2. Escolha o tipo de registro no menu.
3. Preencha os campos.
4. Anexe evidência quando necessário.
5. Clique em Registrar.
6. O ID é criado automaticamente.
7. O Dashboard é atualizado.

## Próximas evoluções recomendadas

- edição/exclusão de registros;
- pesquisa e filtros;
- exportação Excel/PDF;
- gráficos mensais;
- cadastro de usuários/RPAs;
- tela de relatórios;
- controle de permissões;
- backup automático do banco.
