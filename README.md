Observações quanto a adaptações pro GitHub: Modifiquei alguns caminhos e nomes de arquivos que julguei serem desnecessários pro repositório, além de ser uma medida de segurança.

Softwares utilizados: VSCode, DBeaver (cliente SQLite), QGIS.

Funcionalidades:

 O script "comparacao_matricula" serviu para uma função bem específica que é o upload de uma tabela no banco de dados com uma comparação que estava meio difícil de fazer pelo cliente SQL.


<<<<<<< HEAD
 O script "sincronia_auto" provavelmente é o mais importante, ele integra as planilhas preenchidas em tempo real ao banco de dados, basta executar o script ou adicionar ao agendador de tarefas. Além da integração, ele já jogas as planilhas tratadas em formato de tabela e com uma chave em comum para cada endereço em relação a todas tabelas.
=======
 O script "sincronia_auto" provavelmente é o mais importante, ele integra as planilhas preenchidas em tempo real ao banco de dados, basta executar o script ou adicionar ao agendador de tarefas. Além da integração, ele já jogas as planilhas tratadas em formato de tabela.
>>>>>>> 5ec1c5d629dee5aa890d0c13656b764746ff5ac7


 O script "sincronia_CADASTRO_auto" tem uma sacada muito legal, pega a planilha geral da Copasa, com mais de 250K linhas e lança no banco de dados sem as colunas desnecessárias. Além de estar num banco SQLite, que já acelera as consultas, eu tratei as coordenas para WKT, formato aceito no QGIS. A tabela, portanto, está 100% georreferenciada. Ele está à parte por que é um script que mexe com muitos dados e não usaremos ele com frequência.


Observações quanto ao uso de inteligência artificial: A IA foi usada nesse projeto na parte de otimização de alguns códigos, correções eventuais, algumas revisões e na sugestão de algumas funcionalidades que eu não tinha domínio e que foram de grande ajuda (Ex. converter as coordenadas pra WKT, uso do módulo re, etc).

Dúvidas ou sugestões: [LinkedIn](https://www.linkedin.com/in/gustavo-costa-51110223b/) ou no contato do github mesmo.