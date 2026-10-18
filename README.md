**Simulador de Investimentos em Fundos Imobiliários (FIIs)**
Projeto desenvolvido para o desafio Criando uma Ferramenta de Controle de Investimentos com Excel, da DIO, apresentado por Felipe Aguiar (Felipão).

**Objetivo:**
Criar uma ferramenta simples, com aparência de aplicativo, que ajude a simular aportes mensais em fundos imobiliários (FIIs), visualizar a evolução do patrimônio e estimar dividendos. Os valores utilizados são fictícios e apenas educacionais.

**O que a ferramenta responde?**
Na aba APP:
Pergunta	Onde encontrar
Quanto investir por mês?	Célula D10
Por quantos anos investir?	Célula D11
Qual a taxa de rendimento mensal?	Célula D12
Qual o patrimônio acumulado estimado?	Célula D13
Quanto receber em dividendos por mês?	Célula D14


A mesma aba apresenta a sugestão de aporte equivalente a 30% do salário informado, projeções para 2, 5, 10, 20 e 30 anos e uma distribuição ilustrativa do aporte por tipo de FII.
Como usar
1. Abra simulador_fii_modelo_simples.xlsx no Excel.
2. Na aba APP, preencha as células amarelas com salário, aporte, prazo e taxas de exemplo.
3. Consulte o patrimônio acumulado e os dividendos mensais estimados.
4. Compare os cenários de diferentes prazos.
5. No campo Perfil do investidor, selecione Conservador, Moderado ou Agressivo e observe a alteração dos percentuais e valores destinados a cada tipo de fundo.
As células cinzas exibem resultados automáticos e não devem ser alteradas durante o uso normal.
Fórmulas utilizadas
VF (valor futuro) — estima o patrimônio formado com aportes periódicos e rentabilidade mensal constante. No arquivo, a função aparece como FV, nome utilizado internamente na fórmula do Excel:
=FV(D12,D11*12,-D10,0,0)
- D12: taxa mensal;
- D11*12: prazo convertido em meses;
- -D10: aporte mensal, representado como saída de caixa;
- 0: patrimônio inicial;
- 0: aportes realizados ao final de cada período.
Os dividendos mensais estimados são calculados multiplicando o patrimônio projetado pela taxa de rendimento da carteira em D6:
=D13*D6
PROCV — recupera o percentual correspondente à combinação de perfil do investidor e tipo de FII na aba BASE_PERFIS:
=VLOOKUP($C$27&"|"&B31,BASE_PERFIS!$A$2:$D$19,4,FALSE)
Na versão em português do Excel, a função é conhecida como PROCV. A chave composta combina o perfil com o tipo de fundo, por exemplo: Moderado|Papel. O valor distribuído a cada tipo de fundo resulta da multiplicação desse percentual pelo aporte mensal.
As projeções utilizam referências absolutas, como $D$12 e $D$10, para manter fixos os parâmetros ao replicar as fórmulas.
Perfis e percentuais ilustrativos
Os percentuais abaixo foram definidos para fins didáticos e não representam indicação de compra ou alocação financeira.
Tipo de FII	Conservador	Moderado	Agressivo
Papel	40%	32%	20%
Tijolo	35%	35%	25%
Híbridos	10%	8%	15%
FOFs	10%	5%	10%
Desenvolvimento	3%	10%	20%
Hotelarias	2%	10%	10%
Total	100%	100%	100%


A tabela de apoio está na aba BASE_PERFIS e contém a chave composta, o perfil, o tipo de fundo e o percentual.
Estrutura do arquivo
- APP: painel principal de entrada de dados, simulação, cenários e divisão por perfil.
- BASE_PERFIS: tabela auxiliar dos percentuais usados nas buscas.
Personalizações em relação ao modelo do instrutor
Foi adotado um visual simplificado, com duas abas, divisão entre células de entrada e resultados, apresentação direta das cinco perguntas do desafio e percentuais didáticos próprios para os três perfis. O objetivo é facilitar o uso e a compreensão das fórmulas sem um painel excessivamente elaborado.
Intervalos nomeados
Na versão atual do arquivo, as fórmulas utilizam referências de células, como D10 e D12, e referências absolutas. Os intervalos nomeados aporte e taxa_mensal, sugeridos no enunciado, ainda não foram definidos no arquivo. Para implementar essa melhoria no Excel, use Fórmulas > Gerenciador de Nomes e associe:
- aporte → APP!$D$10
- taxa_mensal → APP!$D$12
Depois, as referências correspondentes podem ser substituídas pelos nomes nas fórmulas.
Evidências para a entrega
Após abrir a planilha no Excel, recomenda-se registrar capturas de tela da aba APP com os mesmos valores de aporte, prazo e taxa, alternando somente entre os perfis Conservador e Agressivo. Isso demonstra que a distribuição é atualizada com a seleção do perfil.
As imagens podem ser colocadas na pasta imagens/ do repositório e incorporadas a este README posteriormente.
Observações
Esta ferramenta apresenta simulações matemáticas, não previsões garantidas de retorno. As taxas são constantes no cálculo e não incluem, de forma individualizada, oscilações de cotas, taxas, tributos, inflação ou riscos dos investimentos. O rendimento aplicado à formação do patrimônio e a taxa usada para estimar os dividendos são parâmetros independentes da planilha.
Projeto educacional — DIO | Excel | Análise de Dados.
