# [Diagrama Unifilar Automático](https://www.google.com/?hl=pt-BR)
***Contexto:***
Durante uma disciplina do curso de Engenharia Elétrica, foi proposto o desenvolvimento de uma ferramenta para cálculo de curto-circuito em sistemas elétricos de potência, semelhante ao simulador Anafas. O método adotado utilizava a entrada de dados via planilhas do Excel, contendo o cadastro dos barramentos, linhas de transmissão, transformadores e cargas shunt.

Após a leitura dos elementos e a realização dos cálculos de curto-circuito, tensões nos barramentos e correntes nos trechos, percebi uma limitação importante: a entrada de dados exclusivamente por tabela tornava a visualização do sistema muito abstrata. Isso aumentava a suscetibilidade a erros de digitação e conexões incorretas — principalmente em sistemas com grande número de barramentos e equipamentos conectados.

Com isso, encontrei a biblioteca Schemdraw, em Python, que apresenta uma documentação excelente e permite a criação de circuitos elétricos e eletrônicos, diagramas de blocos, fluxogramas e desenhos técnicos com alta qualidade, exportáveis em diversos formatos (PDF, PNG, JPEG, etc.).

A ideia, então, foi utilizar Python para ler as planilhas de dados do sistema e gerar automaticamente o diagrama unifilar correspondente, de forma organizada e conectada. A maior dificuldade esteve na organização visual dos elementos, especialmente na prevenção de sobreposição ou cruzamentos entre conexões.

Para resolver esse desafio, utilizei conceitos de teoria dos grafos, com o auxílio da biblioteca NetworkX, o que permitiu organizar o circuito de maneira limpa e compreensível. Como a biblioteca Schemdraw não é voltada especificamente para sistemas elétricos de potência, desenvolvi novos elementos personalizados, como símbolos de barramento, transformadores de dois e três enrolamentos e o símbolo de curto-circuito.

Com a base do diagrama pronta, passei à etapa de personalização: colorindo os elementos conforme seus níveis de tensão, adicionando margens e uma legenda — semelhante ao que se vê em desenhos técnicos —, incluindo as principais informações do curto-circuito no barramento em falta.

Além de enriquecer a visualização e facilitar a interpretação dos sistemas simulados, essa ferramenta pode ser especialmente útil quando utilizada em conjunto com simuladores que não apresentam diagrama unifilar integrado, como é o caso do Simulight e do Anatem, proporcionando uma interface visual complementar que aumenta a confiabilidade e compreensão dos resultados.

Por fim, deixo alguns dos diagramas gerados pela aplicação como exemplos e agradeço ao Rafael de Castro Roque e ao Gabriel Bié da Fonseca pelas críticas e sugestões construtivas ao longo do desenvolvimento deste trabalho. Ferramenta desenvolvida dispoível no meu Github e em meu site.

Referências:  
https://schemdraw.readthedocs.io/en/stable/  
https://networkx.org/documentation/stable/  
https://pandas.pydata.org/docs/

---
## Como Usar  
Há diversos parâmetros que podem ser configurados para se obter o diagrama da forma desejada. são eles:  
- Aproximadamente linha 65:  
"posicao_elementos = nx.spring_layout(G, scale=8, iterations=300000, threshold=1e-10)"   

***nx.spring_layout*** modelo de organização de grafos, opções:  
spring_layout → modelo de molas  
circular_layout → em círculo  
shell_layout → camadas concêntricas  
kamada_kawai_layout → distâncias preservadas  
random_layout → posições aleatórias  
spectral_layout → autovalores  
planar_layout → sem cruzamentos (grafo planar)  
spiral_layout → em espiral  
bipartite_layout(G, nodes) → dois grupos (bipartido)  

***scale*** espaço esntre os vértices dos grafos (Elementos), número maior que zero.  

***iterations*** Basicamente número de iterações que o grafo tenta para obter a melhor organização.  


- Aproximadamente linha 71:  
"cores_tensoes = sns.color_palette("tab10", len(niveis_de_tensoes))"  
Essa linha tem como objetivo gerar as cores de forma automática, porém a paleta de cores pode ser alterada.
opções:  
| Paleta         | Palavra-chave            |  
| -------------- | ------------------------ |  
| `"deep"`       | **Padrão elegante**      |  
| `"muted"`      | **Suave**                |  
| `"bright"`     | **Vibrante**             |  
| `"pastel"`     | **Clara/Suave**          |  
| `"dark"`       | **Escura**               |  
| `"colorblind"` | **Acessível**            |  
| `"tab10"`      | **Clássica**             |  
| `"Set1"`       | **Forte**                |  
| `"Set2"`       | **Moderada**             |  
| `"Set3"`       | **Variada**              |  
| `"Paired"`     | **Duplas contrastantes** |  


- Aproximadamente linha 148:  
"posicao_elementos[key][0] *= 8"  
"posicao_elementos[key][1] *= 6"  
Esses numeros podem ser altrados para atingir uma melhor proporção X e Y, a depender se deseja imprimir, tamanho do monitor etc.



- Aproximadamente linha 296:  
"d.config(fontsize=7)"  
Ajustar o tamanho dos textos dos elemntos no diagrama 

<embed src="C:\Users\danie\PycharmProjects\Xlwings\Arquivo_Cadastro_Elementos_Final_Legendado.pdf" width="100%" height="600px" />

