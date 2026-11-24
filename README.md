# Winner Motors — Automação de Pedidos

Aplicação desktop em Python para automatizar o preenchimento de pedidos de **compra e venda** no sistema de pedidos da Winner Motors, utilizando dados de uma planilha Excel.

A interface permite selecionar o tipo de pedido, informar os dados necessários, escolher a planilha e acompanhar a execução por um log.

> **Atenção:** esta automação interage com a tela por teclado e mouse. O sistema de pedidos precisa estar aberto e nas telas esperadas pelo roteiro. Antes de utilizar em operações reais, faça testes com pedidos controlados e confira os dados lançados.

## Funcionalidades

- Leitura de planilhas `.xlsx`.
- Montagem da lista de itens de compra a partir das abas de capacetes e peças.
- Expansão automática dos códigos de capacetes por tamanho (S, M, L e XL), conforme as quantidades informadas na planilha.
- Leitura de itens de venda pela aba `Ajuste de Custo`.
- Preenchimento dos dados do cabeçalho da venda.
- Preenchimento da operação informada na interface.
- Opções de compra para zerar o preço de lista e/ou digitar preço unitário, quando houver preço disponível na leitura da planilha.
- Registro das etapas e dos erros no log da aplicação.
- Botão para solicitar a interrupção da automação.
- Recurso de segurança do PyAutoGUI: levar o ponteiro do mouse ao canto superior esquerdo da tela interrompe a automação.

## Tecnologias e bibliotecas

- Python
- Tkinter / ttk — interface gráfica
- PyAutoGUI — automação de teclado e mouse
- OpenPyXL — leitura de arquivos Excel
- OpenCV — comparação de imagens na identificação visual do aviso
- Pillow — carregamento e manipulação da imagem de referência
- Threading — execução da automação em uma thread separada da interface

## Requisitos

- Windows.
- Python convencional instalado. No ambiente utilizado no desenvolvimento, a execução é feita com o interpretador Python 3.14 convencional (`py -V:3.14`). Evite utilizar a variante *free-threaded* `3.14t`, pois houve incompatibilidade entre essa variante e as bibliotecas binárias NumPy/OpenCV/Pillow.
- Sistema de pedidos aberto e pronto para receber o pedido.
- Planilha Excel `.xlsx` preenchida conforme a estrutura descrita abaixo.
- Arquivo de imagem `aviso_cliente.png` na mesma pasta do script, caso a função de reconhecimento do aviso de crédito esteja habilitada na versão utilizada.

## Instalação

Abra o PowerShell na pasta do projeto e instale as dependências usando o mesmo interpretador que será utilizado para executar o programa:

```powershell
py -V:3.14 -m pip install pyautogui opencv-python openpyxl Pillow
```

Para verificar se as bibliotecas principais estão disponíveis:

```powershell
py -V:3.14 -c "import pyautogui, cv2, numpy, openpyxl, pyscreeze; from PIL import Image; print('Dependências carregadas com sucesso')"
```

Se o comando acima falhar, confira se o `pip` e a execução estão apontando para o mesmo Python. Instalar uma biblioteca em um interpretador e iniciar o script por outro pode causar erros de importação.

## Como executar

No PowerShell, dentro da pasta do projeto:

```powershell
py -V:3.14 pedidos_compra_venda.py
```

Na interface:

1. Escolha `Compra` ou `Venda`.
2. Informe o código da operação.
3. Para vendas, preencha os dados do cabeçalho.
4. Selecione a planilha Excel.
5. Deixe o sistema de pedidos aberto na tela inicial/etapa esperada.
6. Clique em **Iniciar Automação** e confirme as mensagens apresentadas.
7. Acompanhe as etapas pelo **Log de Execução**.

A automação inicia após uma contagem regressiva de cinco segundos. Na rotina de venda descrita no código-base, depois de preencher o cabeçalho o programa aguarda a seleção da primeira linha de produtos no grid. O foco da tela deve estar no local correto para que a sequência de teclas seja aplicada aos campos esperados.

## Estrutura esperada da planilha

O programa lê arquivos `.xlsx`. Os nomes das abas e as colunas utilizadas estão configurados no código.

### Pedido de compra

As abas são lidas a partir da linha 2.

#### Capacetes

Abas reconhecidas:

- `SHARK`
- `KYT`
- `ORIGINE`
- `SUOMY`

| Coluna | Conteúdo |
|---|---|
| D | Código base do capacete, com sufixo separado por ponto |
| E | Descrição |
| F | Quantidade tamanho S |
| G | Quantidade tamanho M |
| H | Quantidade tamanho L |
| I | Quantidade tamanho XL |

O código precisa conter um ponto separando a base do sufixo. Para cada tamanho com quantidade preenchida, o programa gera um código incluindo o tamanho antes do sufixo. Quantidades vazias ou iguais a zero são ignoradas.

Exemplo ilustrativo: um código `ABC.SUF` com quantidades nos tamanhos S e M gera `ABCS.SUF` e `ABCM.SUF`.

#### Peças

Abas reconhecidas:

- `PARTS SHARK`
- `PARTS KYT`
- `PARTS ORIGINE`
- `PARTS SUOMY`

| Coluna | Conteúdo |
|---|---|
| B | Código |
| C | Descrição |
| D | Quantidade |

A aba `DUOTECH` utiliza uma disposição diferente:

| Coluna | Conteúdo |
|---|---|
| C | Código |
| D | Descrição |
| E | Quantidade |

As abas configuradas que não existirem no arquivo são ignoradas. Linhas sem código ou sem quantidade válida não entram na lista.

**Preços em compras:** a leitura atual das abas de compra não possui uma coluna de preço unitário configurada. Por isso, a opção de digitar preço unitário avisa que os campos poderão ficar sem preenchimento. A opção `Zerar preço de lista` preenche esse campo com zero.

### Pedido de venda

A rotina de venda utiliza a aba `Ajuste de Custo`, a partir da linha 6.

| Coluna | Conteúdo |
|---|---|
| F | Código do produto |
| G | Descrição |
| H | Quantidade |
| K | Custo Trade / preço unitário |

O preço de lista é preenchido com `0`. O preço unitário é obtido da coluna K (`Custo Trade`). Se o custo estiver vazio, o programa registra um aviso no log e deixa o campo sem preenchimento.

## Dados do cabeçalho da venda

A interface possui os seguintes campos:

| Campo | Regra |
|---|---|
| Cliente | Numérico, até 6 dígitos; completado com zeros à esquerda |
| Peso bruto | Numérico, até 2 casas decimais |
| Volume 1 | Numérico, até 5 dígitos; completado com zeros à esquerda |
| Transportadora | Numérico, até 6 dígitos; completado com zeros à esquerda |
| Condição de pagamento | Numérico, até 3 dígitos; completado com zeros à esquerda |
| Mensagem p/ nota | Texto livre |
| Operação | Numérica, até 2 dígitos; completada com zero à esquerda |

A forma de pagamento não é solicitada na interface: a rotina preenche o valor fixo `X`. Depois de navegar pelos campos do cabeçalho, o roteiro cola a mensagem da nota e digita o código fixo `008`, conforme o fluxo configurado.

O preenchimento depende da ordem e da quantidade de campos existentes na tela do sistema. Alterações no formulário do sistema podem exigir ajustes na sequência de teclas e nas pausas definidas no script.

## Reconhecimento do aviso de crédito

Na rotina de venda, o sistema pode apresentar um aviso de crédito após a informação do cliente. Como os valores exibidos podem mudar de um cliente para outro, a identificação visual deve utilizar uma parte fixa da janela, e não o texto variável do limite de crédito.

Quando o reconhecimento estiver habilitado:

- Mantenha `aviso_cliente.png` na mesma pasta do arquivo `.py`.
- Use como referência um recorte que contenha somente a parte fixa do aviso (por exemplo, o título `FATV0003 - Dados de crédito:`).
- Não inclua na referência o nome do cliente, o limite de crédito ou outros valores variáveis.
- O programa procura a referência na tela com uma tolerância de similaridade configurada em `confidence`.
- Ao encontrar o aviso, a rotina pressiona `Enter` e continua o fluxo.

O recorte deve ser atualizado caso o título ou o visual fixo da janela seja alterado no sistema. A leitura da imagem pelo Pillow também evita passar diretamente ao OpenCV um caminho de rede que contenha caracteres especiais, como `ç` ou `ã`.

## Ajustes de velocidade

O ritmo da automação é controlado por constantes no início do arquivo:

| Constante | Finalidade |
|---|---|
| `TEMPO_PREPARO` | Contagem regressiva antes de começar |
| `DELAY_ENTRE_ITENS` | Pausa entre os itens |
| `INTERVALO_DIGITACAO` | Intervalo entre caracteres digitados |
| `PAUSA_APOS_DIGITAR` | Pausa depois de preencher um campo |
| `PAUSA_ENTRE_TABS_CABECALHO` | Pausa entre cada tecla `TAB` no cabeçalho |
| `PAUSA_APOS_GRUPO_TABS` | Pausa adicional depois de um grupo de `TAB`s |

Altere esses valores com cautela e teste com poucos itens antes de executar um pedido completo. Aumentar as pausas pode ajudar quando o sistema demora a processar os campos.

## Cuidados e limitações

- Faça uma cópia de segurança da planilha antes de utilizá-la.
- Confira os códigos, quantidades, custos e dados do cabeçalho antes de iniciar.
- Não mova o mouse nem interaja com a janela do sistema durante a digitação automatizada, exceto quando o fluxo solicitar uma ação manual.
- Mantenha a janela do sistema na posição esperada e evite alterar a resolução ou o dimensionamento da tela durante a execução.
- O roteiro utiliza sequências de `TAB`, setas e digitação. Se a ordem dos campos no sistema mudar, o código precisará ser revisado.
- A função de parada solicita a interrupção; dependendo da etapa em andamento, a rotina pode terminar a ação atual antes de parar.
- O PyAutoGUI possui uma proteção de emergência: mover o mouse para o canto superior esquerdo interrompe a automação.
- Teste primeiro em ambiente controlado. O programa não substitui a conferência do pedido lançado no sistema.

## Estrutura sugerida do repositório

```text
pedidos-automacao/
├── pedidos_compra_venda.py
├── aviso_cliente.png          # necessário se o reconhecimento visual estiver habilitado
├── README.md
└── .gitignore
```

Não publique planilhas reais, dados de clientes, credenciais ou informações comerciais internas no repositório. Se a imagem do aviso contiver dados reais, use um recorte anonimizado que mantenha apenas os elementos fixos necessários ao reconhecimento.

## Licença

Defina aqui a licença de distribuição do projeto, caso ele vá ser disponibilizado publicamente. Se o código for destinado apenas ao uso interno da Winner Motors, confirme a autorização da empresa antes de publicar o repositório como público.
