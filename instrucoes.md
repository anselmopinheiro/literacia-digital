# Gerador de Documentos DOCX com Tags

## Visão Geral

Este sistema permite gerar automaticamente múltiplos documentos DOCX a partir de um template base, substituindo tags específicas por dados configuráveis. É especialmente útil para criar documentos de formação, sessões ou qualquer tipo de documento que precise ser personalizado para diferentes turmas, datas ou contextos.

## Características Principais

- ✅ **Preservação completa da formatação** (negrito, itálico, cores, fontes, etc.)
- ✅ **Configuração via JSON** para múltiplas turmas
- ✅ **Processamento automático** de múltiplas sessões
- ✅ **Nomenclatura automática** de ficheiros
- ✅ **Suporte para headers, footers e tabelas**
- ✅ **Interface interativa** para facilitar o uso

## Instalação

### Pré-requisitos

- Python 3.6 ou superior
- Biblioteca `python-docx`

### Instalação da dependência

```bash
pip install python-docx
```

## Estrutura de Ficheiros

```
projeto/
├── script.py                    # Script principal
├── template.docx               # Template base (obrigatório)
├── configuracao_turmas.json    # Configuração das turmas
└── documentos_turmas/          # Pasta de saída (criada automaticamente)
    ├── TurmaA_2024-01-15.docx
    ├── TurmaA_2024-01-22.docx
    └── ...
```

## Tags Suportadas

O sistema substitui as seguintes tags no documento template:

| Tag | Descrição | Exemplo |
|-----|-----------|---------|
| `<<Turma>>` | Nome da turma | TurmaA |
| `<<DT>>` | Direção técnica | Direção Técnica A |
| `<<ronda>>` | Identificação da ronda | 1ª Ronda |
| `<<sessao>>` | Número da sessão | 1, 2, 3... |
| `<<data>>` | Data da sessão | 15/01/2024 |
| `<<Docente1>>` | Primeiro docente | Prof. João Silva |
| `<<Docente2>>` | Segundo docente | Prof. Maria Santos |
| `<<Docente3>>` | Terceiro docente | Prof. Carlos Pereira |
| `<<Docente4>>` | Quarto docente | Prof. Ana Costa |

## Configuração JSON

### Estrutura do Ficheiro

O ficheiro `configuracao_turmas.json` segue esta estrutura:

```json
{
    "turmas": {
        "TurmaA": {
            "nome": "TurmaA",
            "dt": "Direção Técnica A",
            "ronda": "1ª Ronda",
            "docentes": {
                "docente1": "Prof. João Silva",
                "docente2": "Prof. Maria Santos",
                "docente3": "Prof. Carlos Pereira",
                "docente4": "Prof. Ana Costa"
            },
            "sessoes": [
                {"sessao": 1, "data": "2024-01-15"},
                {"sessao": 2, "data": "2024-01-22"},
                {"sessao": 3, "data": "2024-01-29"}
            ]
        },
        "TurmaB": {
            "nome": "TurmaB",
            "dt": "Direção Técnica B",
            "ronda": "2ª Ronda",
            "docentes": {
                "docente1": "Prof. Ricardo Oliveira",
                "docente2": "Prof. Luísa Fernandes",
                "docente3": "Prof. Miguel Torres",
                "docente4": "Prof. Sofia Ribeiro"
            },
            "sessoes": [
                {"sessao": 1, "data": "2024-02-12"},
                {"sessao": 2, "data": "2024-02-19"}
            ]
        }
    }
}
```

### Personalização

#### Adicionar Nova Turma

```json
"TurmaD": {
    "nome": "TurmaD",
    "dt": "Direção Técnica D",
    "ronda": "4ª Ronda",
    "docentes": {
        "docente1": "Prof. Nome1",
        "docente2": "Prof. Nome2",
        "docente3": "Prof. Nome3",
        "docente4": "Prof. Nome4"
    },
    "sessoes": [
        {"sessao": 1, "data": "2024-04-01"},
        {"sessao": 2, "data": "2024-04-08"}
    ]
}
```

#### Formato de Datas

- **Formato obrigatório**: `YYYY-MM-DD`
- **Exemplos válidos**: `2024-01-15`, `2024-12-25`
- **Conversão automática**: O sistema converte para `DD/MM/YYYY` no documento final

## Como Usar

### Passo 1: Preparar o Template

1. Crie um documento Word (`template.docx`) com o conteúdo base
2. Insira as tags onde pretende que os dados sejam substituídos
3. Aplique toda a formatação desejada (negrito, cores, etc.)

**Exemplo de template:**
```
Turma: <<Turma>>
Direção: <<DT>>
Ronda: <<ronda>>
Sessão nº: <<sessao>>
Data: <<data>>

Docentes:
- <<Docente1>>
- <<Docente2>>
- <<Docente3>>
- <<Docente4>>
```

### Passo 2: Executar o Script

```bash
python script.py
```

### Passo 3: Escolher Opção

```
=== GERADOR DE DOCUMENTOS DOCX COM JSON ===
1. Criar ficheiro de configuração JSON
2. Processar todas as turmas do JSON
3. Mostrar estrutura do JSON

Escolha uma opção (1-3):
```

#### Opção 1 - Primeira Execução
- Cria o ficheiro `configuracao_turmas.json` com dados de exemplo
- Permite personalizar as configurações

#### Opção 2 - Gerar Documentos
- Processa todas as turmas configuradas no JSON
- Cria os documentos na pasta `documentos_turmas/`

#### Opção 3 - Ajuda
- Mostra a estrutura detalhada do JSON
- Fornece exemplos de configuração

## Nomenclatura dos Ficheiros

Os ficheiros gerados seguem o padrão: **`{turma}_{data}.docx`**

**Exemplos:**
- `TurmaA_2024-01-15.docx`
- `TurmaB_2024-02-12.docx`
- `TurmaC_2024-03-19.docx`

## Tratamento de Erros

### Erros Comuns e Soluções

#### ❌ "Template não encontrado"
**Solução:** Certifique-se que o ficheiro `template.docx` está no mesmo diretório do script.

#### ❌ "JSON não encontrado"
**Solução:** Execute primeiro a Opção 1 para criar o ficheiro de configuração.

#### ❌ "Data inválida"
**Solução:** Verifique se as datas no JSON estão no formato `YYYY-MM-DD`.

#### ❌ "Erro de codificação"
**Solução:** Certifique-se que o JSON está guardado com codificação UTF-8.

## Funcionalidades Avançadas

### Preservação de Formatação

O sistema preserva toda a formatação aplicada às tags no template:

- **Texto**: negrito, itálico, sublinhado, riscado
- **Fontes**: tipo, tamanho, cor
- **Parágrafos**: alinhamento, espaçamento, indentação
- **Fundo**: cores de realce
- **Estilos**: estilos aplicados do Word

### Suporte Completo a Elementos Word

- ✅ Parágrafos normais
- ✅ Tabelas (todas as células)
- ✅ Headers (cabeçalhos)
- ✅ Footers (rodapés)
- ✅ Caixas de texto
- ✅ Notas de rodapé

### Processamento Inteligente

O sistema divide o texto em "runs" (segmentos com formatação uniforme) e:
1. Localiza exatamente onde cada tag está posicionada
2. Substitui apenas o conteúdo da tag
3. Mantém toda a formatação original
4. Funciona mesmo quando tags estão divididas entre múltiplos runs

## Casos de Uso

### Educação/Formação
- Documentos de sessões de formação
- Fichas de avaliação por turma
- Calendários de aulas personalizados

### Administração
- Convocatórias personalizadas
- Relatórios por departamento
- Documentos oficiais com dados variáveis

### Eventos
- Certificados personalizados
- Programas de eventos
- Listas de participantes

## Limitações

- **Formato de data**: Apenas aceita `YYYY-MM-DD` no JSON
- **Encoding**: Ficheiros devem estar em UTF-8
- **Template**: Deve ser um ficheiro `.docx` válido
- **Tags**: Devem usar exatamente o formato `<<nome>>`

## Exemplo Completo

### 1. Template (template.docx)
```
SESSÃO DE FORMAÇÃO

Turma: <<Turma>>
Direção Técnica: <<DT>>
Ronda: <<ronda>>
Sessão: <<sessao>>
Data: <<data>>

EQUIPA FORMATIVA:
Docente 1: <<Docente1>>
Docente 2: <<Docente2>>
Docente 3: <<Docente3>>
Docente 4: <<Docente4>>
```

### 2. Configuração JSON
```json
{
    "turmas": {
        "Informatica": {
            "nome": "Informática",
            "dt": "Direção Técnica IT",
            "ronda": "1ª Ronda",
            "docentes": {
                "docente1": "Prof. Ana Silva",
                "docente2": "Prof. João Costa",
                "docente3": "Prof. Maria Pereira",
                "docente4": "Prof. Carlos Santos"
            },
            "sessoes": [
                {"sessao": 1, "data": "2024-01-15"},
                {"sessao": 2, "data": "2024-01-22"}
            ]
        }
    }
}
```

### 3. Resultado (Informatica_2024-01-15.docx)
```
SESSÃO DE FORMAÇÃO

Turma: Informática
Direção Técnica: Direção Técnica IT
Ronda: 1ª Ronda
Sessão: 1
Data: 15/01/2024

EQUIPA FORMATIVA:
Docente 1: Prof. Ana Silva
Docente 2: Prof. João Costa
Docente 3: Prof. Maria Pereira
Docente 4: Prof. Carlos Santos
```

## Suporte e Manutenção

### Versão Atual
- **Versão**: 2.0
- **Data**: 2025
- **Python**: 3.6+



### Contacto
Para suporte ou sugestões, consulte a documentação técnica do código Python incluída no script.

---

**Nota**: Este sistema foi desenvolvido para maximizar a eficiência na criação de documentos personalizados, mantendo sempre a qualidade e formatação profissional dos documentos originais.