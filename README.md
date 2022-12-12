# -Colocar o arquivo excel na pasta "file" na raiz da instalação

### ./file/Seu arquivo para importação.xlsx

#### Habilitando a execução assíncrona do DJANGO


```python
import os

# Para funcionamento do 'for' na hora da inclusão no banco, valor padrão é False
os.environ["DJANGO_ALLOW_ASYNC_UNSAFE"] = "True"
```

# -Importação do arquivo e Inserção na base de dados


```python
import openpyxl
from polls.models import Choice, Question
from django.utils import timezone

path = openpyxl.load_workbook(r"./file/Planilha de perguntas e respostas.xlsx")
lista = path['Planilha1']

# Itera sobre as linahs da planilha em Excel e salva no banco de dados as perguntas, escolhas e votos de cada uma
for linha in range(2, lista.max_row+1):
    question = lista.cell(linha, 1).value
#   Executado por primeiro para criação da Question  
    q = Question(question_text=question, pub_date=timezone.now())
    q.save()
    choice1 = lista.cell(linha, 2).value
    votes1 = lista.cell(linha, 3).value
    choice2 = lista.cell(linha, 4).value
    votes2 = lista.cell(linha, 5).value
    choice3 = lista.cell(linha, 6).value
    votes3 = lista.cell(linha, 7).value
    choice4 = lista.cell(linha, 8).value
    votes4 = lista.cell(linha, 9).value

#   Criando as choice e votes de cada Question
    q = Question.objects.get(question_text=question)
    q.choice_set.create(choice_text=choice1, votes=votes1)
    q.choice_set.create(choice_text=choice2, votes=votes2)
    q.choice_set.create(choice_text=choice3, votes=votes3)
    q.choice_set.create(choice_text=choice4, votes=votes4)
    q.save()
```


```python

```
