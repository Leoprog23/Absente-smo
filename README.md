O sistema foi construído em Python, utilizando ferramentas como SQLite, Pandas, Tkinter e o Pyodbc.
O intuito dele é fazer uma gestão via sistema no tocante a absenteísmo de funcionários, o sistema é alimentado por um arquivo EXCEL e no mesmo encontra-se algumas informações de cada colaborador do setor em questão (Logística).
As informações de cada funcionário são utilizadas para atualizações dinâmicas de acordo com os campos do sistema, armazenados em um banco de dados SQLite e renderizados em uma tabela, dentro do próprio sistema essa tabela já é exibida.
A tabela possui filtros de acordo com cada coluna criada no BD, botões iterativos (Exportar para Excel, Limpar os filtros da tabela e atualização do BD iterando com a tabela exibida)