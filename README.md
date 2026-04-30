OBSERVAÇÕES QUANTO À FIDELIDADE DOS DADOS: Modifiquei caminhos e nomes de arquivos, bancos e tabelas que julguei serem desnecessários pro repositório, além de ser uma medida de segurança. Estrutura é universal.

Softwares utilizados: VSCode, DBeaver (cliente SQLite), QGIS.

Funcionalidades:

 O script "comparacao_matricula" serviu para uma função bem específica que é o upload de uma tabela no banco de dados com uma comparação que estava meio difícil de fazer pelo cliente SQL.

 O script "sincronia_auto" provavelmente é o mais importante, ele integra as planilhas preenchidas em tempo real ao banco de dados, basta executar o script ou adicionar ao agendador de tarefas. Além da integração, ele já jogas as planilhas tratadas em formato de tabela e com uma chave em comum para cada endereço em relação a todas tabelas.

 O script "sincronia_CADASTRO_auto" tem uma sacada muito legal, pega a planilha geral da Copasa, com mais de 250K linhas e lança no banco de dados sem as colunas desnecessárias. Além de estar num banco SQLite, que já acelera as consultas, eu tratei as coordenas para WKT, formato aceito no QGIS. A tabela, portanto, está 100% georreferenciada. Ele está à parte por que é um script que mexe com muitos dados e não usaremos ele com frequência.

O script "resolvendo_irregularidades" usa de fuzzy matching pra tratar de erros de despadronização e erros de digitação. Exemplo: "RUAA|NUMERO2|VILALUGAR" é o mesmo endereço de de "RUAA|NUMERO2|LUGAR", porém o computador lê como informações completamente diferentes. O fuzzy matching usa de comparação e porcentagem para determinar se é erro de digitação ou informações totalmente divergentes, de fato. Este script de auditoria pega as diferenças de digitação e atribui o valor padrão correto pro valor padrão incorreto no banco de dados. 
Além disso, usei estratégias para evitar SQL injections. Apesar de ser improvável no meu ambiente de trabalho, foi uma forma de estudo e de conhecer essas boas práticas.

Observações quanto ao uso de inteligência artificial: A IA foi usada nesse projeto na parte de otimização de alguns códigos, correções eventuais, algumas revisões e na sugestão de algumas funcionalidades que eu não tinha domínio e que foram de grande ajuda (Ex. converter as coordenadas pra WKT, uso do módulo re, etc).

Dúvidas ou sugestões: [LinkedIn](linkedin.com/in/gustavo-costa-comp/) ou no contato do github mesmo.
