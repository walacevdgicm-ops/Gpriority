# ✅ GPriority — Gestor de Tarefas e Prioridades

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/UI-CustomTkinter-1f6feb)
![SQLite](https://img.shields.io/badge/SQLite-003B57?logo=sqlite&logoColor=white)

Aplicativo desktop moderno para **organizar tarefas por prioridade**, com dashboard, calendário mensal, notificações de prazos e exportação de relatórios em **Excel e PDF**.

## ✨ Funcionalidades

- 📊 **Dashboard** com cards de resumo e gráficos por status e por prioridade
- 📝 Criar, editar, concluir e excluir tarefas
- ⭐ Marcar como **favorita** e 📌 **fixar** tarefas importantes no topo
- 🏷️ **Categorias**: Trabalho, Estudos, Casa, Financeiro, Pessoal
- 🔎 **Busca e filtros** por texto, prioridade, categoria, data e favoritas, com ordenação
- 📅 **Calendário mensal** em português mostrando as tarefas de cada dia
- 🔔 **Notificações** de tarefas vencidas e que vencem amanhã ao abrir o app
- 📤 **Exportação** para **Excel (.xlsx)** formatado e **PDF**, além de impressão
- 💾 Tudo salvo automaticamente em **SQLite**

## 🛠️ Tecnologias

`Python` · `CustomTkinter` · `SQLite` · `Matplotlib` · `openpyxl` · `tkcalendar`

## 🚀 Como executar

```bash
git clone https://github.com/walacevdgicm-ops/gpriority.git
cd gpriority
pip install -r requirements.txt
python gpriority.py
```

O banco `gpriority.db` é criado ao lado do script, e os relatórios vão para a pasta `exportacoes/`.

## 🛣️ Próximos passos

- [ ] Tarefas recorrentes (diárias, semanais, mensais)
- [ ] Sincronização com Google Agenda
- [ ] Tema claro/escuro configurável

## 👨‍💻 Autor

**Walace Soares Vidigal Junior** · [GitHub](https://github.com/walacevdgicm-ops)

Licença [MIT](LICENSE)
