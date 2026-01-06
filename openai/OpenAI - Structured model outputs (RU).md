---
id: 0000-0000-000A
name: development-pipeline
title: Development pipeline
description: Шаблон документа с пайплайном разработки проекта
tags:
  - docs
  - development-pipeline
  - template
created_at: 2025-09-23T07:44:00
updated_at: 2025-09-23T07:45:00
version: 0.01
language: ru
---
Структурированные выходные данные модели (Structured model outputs)
========================

Гарантия того, что текстовые ответы модели соответствуют определяемой вами JSON-схеме.

JSON является одним из наиболее широко используемых форматов в мире для обмена данными между приложениями.

Structured Outputs — это функция, которая гарантирует, что модель всегда будет генерировать ответы, соответствующие предоставленной вами [JSON Schema](https://json-schema.org/overview/what-is-jsonschema), так что вам не нужно беспокоиться о том, что модель пропустит обязательный ключ или "галлюцинирует" недопустимое значение перечисления.

Некоторые преимущества Structured Outputs:

1.  **Надежная типобезопасность:** Нет необходимости проверять или повторять запросы из-за некорректно отформатированных ответов.
2.  **Явные отказы:** Отказы модели на основе соображений безопасности теперь можно программно обнаружить.
3.  **Более простое составление промптов:** Не нужно использовать строгие формулировки в промптах для достижения согласованного форматирования.

В дополнение к поддержке JSON Schema в REST API, SDK OpenAI для [Python](https://github.com/openai/openai-python/blob/main/helpers.md#structured-outputs-parsing-helpers) и [JavaScript](https://github.com/openai/openai-node/blob/master/helpers.md#structured-outputs-parsing-helpers) также упрощают определение схем объектов с использованием [Pydantic](https://docs.pydantic.dev/latest/) и [Zod](https://zod.dev/) соответственно. Ниже вы можете увидеть, как извлекать информацию из неструктурированного текста, соответствующего схеме, определенной в коде.

Получение структурированного ответа

```javascript
import OpenAI from "openai";
import { zodTextFormat } from "openai/helpers/zod";
import { z } from "zod";

const openai = new OpenAI();

const CalendarEvent = z.object({
  name: z.string(),
  date: z.string(),
  participants: z.array(z.string()),
});

const response = await openai.responses.parse({
  model: "gpt-4o-2024-08-06",
  input: [
    { role: "system", content: "Извлеките информацию о событии." },
    {
      role: "user",
      content: "Алиса и Боб собираются на научную ярмарку в пятницу.",
    },
  ],
  text: {
    format: zodTextFormat(CalendarEvent, "event"),
  },
});

const event = response.output_parsed;
```

```python
from openai import OpenAI
from pydantic import BaseModel

client = OpenAI()

class CalendarEvent(BaseModel):
    name: str
    date: str
    participants: list[str]

response = client.responses.parse(
    model="gpt-4o-2024-08-06",
    input=[
        {"role": "system", "content": "Извлеките информацию о событии."},
        {
            "role": "user",
            "content": "Алиса и Боб собираются на научную ярмарку в пятницу.",
        },
    ],
    text_format=CalendarEvent,
)

event = response.output_parsed
```

### Поддерживаемые модели

Structured Outputs доступна в наших [последних больших языковых моделях](/docs/models), начиная с GPT-4o. Более старые модели, такие как `gpt-4-turbo` и более ранние, могут использовать вместо этого [JSON mode](/docs/guides/structured-outputs#json-mode).

Когда использовать Structured Outputs через вызов функций (function calling) vs через text.format

--------------------------------------------------------------------------

Structured Outputs доступна в двух формах в OpenAI API:

1.  При использовании [вызова функций (function calling)](/docs/guides/function-calling)
2.  При использовании формата ответа `json_schema`

Вызов функций полезен, когда вы создаете приложение, которое связывает модели и функциональность вашего приложения.

Например, вы можете дать модели доступ к функциям, которые запрашивают базу данных, чтобы создать AI-ассистента, который может помогать пользователям с их заказами, или к функциям, которые могут взаимодействовать с пользовательским интерфейсом.

И наоборот, Structured Outputs через `response_format` более подходят, когда вы хотите указать структурированную схему для использования, когда модель отвечает пользователю, а не когда модель вызывает инструмент.

Например, если вы создаете приложение-репетитор по математике, вы можете захотеть, чтобы ассистент отвечал вашему пользователю, используя определенную JSON-схему, чтобы вы могли генерировать пользовательский интерфейс, который отображает различные части вывода модели distinctными способами.

Проще говоря:

*   Если вы подключаете модель к инструментам, функциям, данным и т.д. в вашей системе, то вам следует использовать вызов функций.
*   Если вы хотите структурировать вывод модели, когда она отвечает пользователю, то вам следует использовать структурированный `text.format`.

Оставшаяся часть этого руководства будет посвящена случаям использования, не связанным с вызовом функций, в Responses API. Чтобы узнать больше об использовании Structured Outputs с вызовом функций, ознакомьтесь с руководством по [вызову функций](/docs/guides/function-calling#function-calling-with-structured-outputs).

### Structured Outputs vs JSON mode

Structured Outputs — это эволюция [JSON mode](/docs/guides/structured-outputs#json-mode). В то время как оба гарантируют генерацию валидного JSON, только Structured Outputs гарантирует соответствие схеме. И Structured Outputs, и JSON mode поддерживаются в Responses API, Chat Completions API, Assistants API, Fine-tuning API и Batch API.

Мы рекомендуем всегда использовать Structured Outputs вместо JSON mode, когда это возможно.

Однако Structured Outputs с `response_format: {type: "json_schema", ...}` поддерживается только снимками моделей `gpt-4o-mini`, `gpt-4o-mini-2024-07-18`, `gpt-4o-2024-08-06` и более поздними.

| |Structured Outputs|JSON Mode|
|---|---|---|
|Выводит валидный JSON|Да|Да|
|Соответствует схеме|Да (см. поддерживаемые схемы)|Нет|
|Совместимые модели|gpt-4o-mini, gpt-4o-2024-08-06 и новее|gpt-3.5-turbo, gpt-4-* и gpt-4o-* модели|
|Включение|`text: { format: { type: "json_schema", "strict": true, "schema": ... } }`|`text: { format: { type: "json_object" } }`|

Примеры
--------

Цепочка рассуждений (Chain of thought)

### Цепочка рассуждений (Chain of thought)

Вы можете попросить модель выводить ответ структурированно, шаг за шагом, чтобы провести пользователя через решение.

Structured Outputs для репетиторства по математике с цепочкой рассуждений

```javascript
import OpenAI from "openai";
import { zodTextFormat } from "openai/helpers/zod";
import { z } from "zod";

const openai = new OpenAI();

const Step = z.object({
  explanation: z.string(),
  output: z.string(),
});

const MathReasoning = z.object({
  steps: z.array(Step),
  final_answer: z.string(),
});

const response = await openai.responses.parse({
  model: "gpt-4o-2024-08-06",
  input: [
    {
      role: "system",
      content:
        "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом.",
    },
    { role: "user", content: "как решить 8x + 7 = -23" },
  ],
  text: {
    format: zodTextFormat(MathReasoning, "math_reasoning"),
  },
});

const math_reasoning = response.output_parsed;
```

```python
from openai import OpenAI
from pydantic import BaseModel

client = OpenAI()

class Step(BaseModel):
    explanation: str
    output: str

class MathReasoning(BaseModel):
    steps: list[Step]
    final_answer: str

response = client.responses.parse(
    model="gpt-4o-2024-08-06",
    input=[
        {
            "role": "system",
            "content": "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом.",
        },
        {"role": "user", "content": "как решить 8x + 7 = -23"},
    ],
    text_format=MathReasoning,
)

math_reasoning = response.output_parsed
```

```bash
curl https://api.openai.com/v1/responses \
  -H "Authorization: Bearer $OPENAI_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "model": "gpt-4o-2024-08-06",
    "input": [
      {
        "role": "system",
        "content": "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом."
      },
      {
        "role": "user",
        "content": "как решить 8x + 7 = -23"
      }
    ],
    "text": {
      "format": {
        "type": "json_schema",
        "name": "math_reasoning",
        "schema": {
          "type": "object",
          "properties": {
            "steps": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "explanation": { "type": "string" },
                  "output": { "type": "string" }
                },
                "required": ["explanation", "output"],
                "additionalProperties": false
              }
            },
            "final_answer": { "type": "string" }
          },
          "required": ["steps", "final_answer"],
          "additionalProperties": false
        },
        "strict": true
      }
    }
  }'
```

#### Пример ответа

```json
{
  "steps": [
    {
      "explanation": "Начнем с уравнения 8x + 7 = -23.",
      "output": "8x + 7 = -23"
    },
    {
      "explanation": "Вычтем 7 из обеих частей, чтобы изолировать член с переменной.",
      "output": "8x = -23 - 7"
    },
    {
      "explanation": "Упростим правую часть уравнения.",
      "output": "8x = -30"
    },
    {
      "explanation": "Разделим обе части на 8, чтобы найти x.",
      "output": "x = -30 / 8"
    },
    {
      "explanation": "Упростим дробь.",
      "output": "x = -15 / 4"
    }
  ],
  "final_answer": "x = -15 / 4"
}
```

Структурированное извлечение данных

### Структурированное извлечение данных

Вы можете определить структурированные поля для извлечения из неструктурированных входных данных, таких как научные статьи.

Извлечение данных из научных статей с использованием Structured Outputs

```javascript
import OpenAI from "openai";
import { zodTextFormat } from "openai/helpers/zod";
import { z } from "zod";

const openai = new OpenAI();

const ResearchPaperExtraction = z.object({
  title: z.string(),
  authors: z.array(z.string()),
  abstract: z.string(),
  keywords: z.array(z.string()),
});

const response = await openai.responses.parse({
  model: "gpt-4o-2024-08-06",
  input: [
    {
      role: "system",
      content:
        "Вы — эксперт по структурированному извлечению данных. Вам будет предоставлен неструктурированный текст из научной статьи, и вы должны преобразовать его в заданную структуру.",
    },
    { role: "user", content: "..." },
  ],
  text: {
    format: zodTextFormat(ResearchPaperExtraction, "research_paper_extraction"),
  },
});

const research_paper = response.output_parsed;
```

```python
from openai import OpenAI
from pydantic import BaseModel

client = OpenAI()

class ResearchPaperExtraction(BaseModel):
    title: str
    authors: list[str]
    abstract: str
    keywords: list[str]

response = client.responses.parse(
    model="gpt-4o-2024-08-06",
    input=[
        {
            "role": "system",
            "content": "Вы — эксперт по структурированному извлечению данных. Вам будет предоставлен неструктурированный текст из научной статьи, и вы должны преобразовать его в заданную структуру.",
        },
        {"role": "user", "content": "..."},
    ],
    text_format=ResearchPaperExtraction,
)

research_paper = response.output_parsed
```

```bash
curl https://api.openai.com/v1/responses \
  -H "Authorization: Bearer $OPENAI_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "model": "gpt-4o-2024-08-06",
    "input": [
      {
        "role": "system",
        "content": "Вы — эксперт по структурированному извлечению данных. Вам будет предоставлен неструктурированный текст из научной статьи, и вы должны преобразовать его в заданную структуру."
      },
      {
        "role": "user",
        "content": "..."
      }
    ],
    "text": {
      "format": {
        "type": "json_schema",
        "name": "research_paper_extraction",
        "schema": {
          "type": "object",
          "properties": {
            "title": { "type": "string" },
            "authors": {
              "type": "array",
              "items": { "type": "string" }
            },
            "abstract": { "type": "string" },
            "keywords": {
              "type": "array",
              "items": { "type": "string" }
            }
          },
          "required": ["title", "authors", "abstract", "keywords"],
          "additionalProperties": false
        },
        "strict": true
      }
    }
  }'
```

#### Пример ответа

```json
{
  "title": "Применение квантовых алгоритмов в межзвездной навигации: новый рубеж",
  "authors": [
    "Д-р Стелла Вояжер",
    "Д-р Нова Стар",
    "Д-р Лайра Хантер"
  ],
  "abstract": "В этой статье исследуется использование квантовых алгоритмов для улучшения систем межзвездной навигации. Используя квантовую суперпозицию и запутанность, предлагаемая навигационная система может вычислять оптимальные пути путешествия через пространственно-временные аномалии более эффективно, чем классические методы. Экспериментальные симуляции предполагают значительное сокращение времени путешествия и расхода топлива для межзвездных миссий.",
  "keywords": [
    "Квантовые алгоритмы",
    "межзвездная навигация",
    "пространственно-временные аномалии",
    "квантовая суперпозиция",
    "квантовая запутанность",
    "космические путешествия"
  ]
}
```

Генерация пользовательского интерфейса (UI)

### Генерация пользовательского интерфейса (UI)

Вы можете генерировать валидный HTML, представляя его в виде рекурсивных структур данных с ограничениями, такими как перечисления (enums).

Генерация HTML с использованием Structured Outputs

```javascript
import OpenAI from "openai";
import { zodTextFormat } from "openai/helpers/zod";
import { z } from "zod";

const openai = new OpenAI();

const UI = z.lazy(() =>
  z.object({
    type: z.enum(["div", "button", "header", "section", "field", "form"]),
    label: z.string(),
    children: z.array(UI),
    attributes: z.array(
      z.object({
        name: z.string(),
        value: z.string(),
      })
    ),
  })
);

const response = await openai.responses.parse({
  model: "gpt-4o-2024-08-06",
  input: [
    {
      role: "system",
      content: "Вы — AI-генератор пользовательского интерфейса. Преобразуйте входные данные пользователя в UI.",
    },
    {
      role: "user",
      content: "Создайте форму профиля пользователя",
    },
  ],
  text: {
    format: zodTextFormat(UI, "ui"),
  },
});

const ui = response.output_parsed;
```

```python
from enum import Enum
from typing import List

from openai import OpenAI
from pydantic import BaseModel

client = OpenAI()

class UIType(str, Enum):
    div = "div"
    button = "button"
    header = "header"
    section = "section"
    field = "field"
    form = "form"

class Attribute(BaseModel):
    name: str
    value: str

class UI(BaseModel):
    type: UIType
    label: str
    children: List["UI"]
    attributes: List[Attribute]

UI.model_rebuild()  # Это требуется для включения рекурсивных типов

class Response(BaseModel):
    ui: UI

response = client.responses.parse(
    model="gpt-4o-2024-08-06",
    input=[
        {
            "role": "system",
            "content": "Вы — AI-генератор пользовательского интерфейса. Преобразуйте входные данные пользователя в UI.",
        },
        {"role": "user", "content": "Создайте форму профиля пользователя"},
    ],
    text_format=Response,
)

ui = response.output_parsed
```

```bash
curl https://api.openai.com/v1/responses \
  -H "Authorization: Bearer $OPENAI_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "model": "gpt-4o-2024-08-06",
    "input": [
      {
        "role": "system",
        "content": "Вы — AI-генератор пользовательского интерфейса. Преобразуйте входные данные пользователя в UI."
      },
      {
        "role": "user",
        "content": "Создайте форму профиля пользователя"
      }
    ],
    "text": {
      "format": {
        "type": "json_schema",
        "name": "ui",
        "description": "Динамически сгенерированный пользовательский интерфейс",
        "schema": {
          "type": "object",
          "properties": {
            "type": {
              "type": "string",
              "description": "Тип компонента UI",
              "enum": ["div", "button", "header", "section", "field", "form"]
            },
            "label": {
              "type": "string",
              "description": "Метка компонента UI, используется для кнопок или полей формы"
            },
            "children": {
              "type": "array",
              "description": "Вложенные компоненты UI",
              "items": {"$ref": "#"}
            },
            "attributes": {
              "type": "array",
              "description": "Произвольные атрибуты для компонента UI, подходящие для любого элемента",
              "items": {
                "type": "object",
                "properties": {
                  "name": {
                    "type": "string",
                    "description": "Имя атрибута, например, onClick или className"
                  },
                  "value": {
                    "type": "string",
                    "description": "Значение атрибута"
                  }
                },
                "required": ["name", "value"],
                "additionalProperties": false
              }
            }
          },
          "required": ["type", "label", "children", "attributes"],
          "additionalProperties": false
        },
        "strict": true
      }
    }
  }'
```

#### Пример ответа

```json
{
  "type": "form",
  "label": "Форма профиля пользователя",
  "children": [
    {
      "type": "div",
      "label": "",
      "children": [
        {
          "type": "field",
          "label": "Имя",
          "children": [],
          "attributes": [
            {
              "name": "type",
              "value": "text"
            },
            {
              "name": "name",
              "value": "firstName"
            },
            {
              "name": "placeholder",
              "value": "Введите ваше имя"
            }
          ]
        },
        {
          "type": "field",
          "label": "Фамилия",
          "children": [],
          "attributes": [
            {
              "name": "type",
              "value": "text"
            },
            {
              "name": "name",
              "value": "lastName"
            },
            {
              "name": "placeholder",
              "value": "Введите вашу фамилию"
            }
          ]
        }
      ],
      "attributes": []
    },
    {
      "type": "button",
      "label": "Отправить",
      "children": [],
      "attributes": [
        {
          "name": "type",
          "value": "submit"
        }
      ]
    }
  ],
  "attributes": [
    {
      "name": "method",
      "value": "post"
    },
    {
      "name": "action",
      "value": "/submit-profile"
    }
  ]
}
```

Модерация

### Модерация

Вы можете классифицировать входные данные по нескольким категориям, что является распространенным способом модерации.

Модерация с использованием Structured Outputs

```javascript
import OpenAI from "openai";
import { zodTextFormat } from "openai/helpers/zod";
import { z } from "zod";

const openai = new OpenAI();

const ContentCompliance = z.object({
  is_violating: z.boolean(),
  category: z.enum(["violence", "sexual", "self_harm"]).nullable(),
  explanation_if_violating: z.string().nullable(),
});

const response = await openai.responses.parse({
    model: "gpt-4o-2024-08-06",
    input: [
      {
        "role": "system",
        "content": "Определите, нарушает ли ввод пользователя определенные правила, и объясните, если нарушает."
      },
      {
        "role": "user",
        "content": "Как подготовиться к собеседованию на работу?"
      }
    ],
    text: {
        format: zodTextFormat(ContentCompliance, "content_compliance"),
    },
});

const compliance = response.output_parsed;
```

```python
from enum import Enum
from typing import Optional

from openai import OpenAI
from pydantic import BaseModel

client = OpenAI()

class Category(str, Enum):
    violence = "violence"
    sexual = "sexual"
    self_harm = "self_harm"

class ContentCompliance(BaseModel):
    is_violating: bool
    category: Optional[Category]
    explanation_if_violating: Optional[str]

response = client.responses.parse(
    model="gpt-4o-2024-08-06",
    input=[
        {
            "role": "system",
            "content": "Определите, нарушает ли ввод пользователя определенные правила, и объясните, если нарушает.",
        },
        {"role": "user", "content": "Как подготовиться к собеседованию на работу?"},
    ],
    text_format=ContentCompliance,
)

compliance = response.output_parsed
```

```bash
curl https://api.openai.com/v1/responses \
  -H "Authorization: Bearer $OPENAI_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "model": "gpt-4o-2024-08-06",
    "input": [
      {
        "role": "system",
        "content": "Определите, нарушает ли ввод пользователя определенные правила, и объясните, если нарушает."
      },
      {
        "role": "user",
        "content": "Как подготовиться к собеседованию на работу?"
      }
    ],
    "text": {
      "format": {
        "type": "json_schema",
        "name": "content_compliance",
        "description": "Определяет, нарушает ли контент определенные правила модерации",
        "schema": {
          "type": "object",
          "properties": {
            "is_violating": {
              "type": "boolean",
              "description": "Указывает, нарушает ли контент правила"
            },
            "category": {
              "type": ["string", "null"],
              "description": "Тип нарушения, если контент нарушает правила. В противном случае - null.",
              "enum": ["violence", "sexual", "self_harm"]
            },
            "explanation_if_violating": {
              "type": ["string", "null"],
              "description": "Объяснение, почему контент является нарушающим"
            }
          },
          "required": ["is_violating", "category", "explanation_if_violating"],
          "additionalProperties": false
        },
        "strict": true
      }
    }
  }'
```

#### Пример ответа

```json
{
  "is_violating": false,
  "category": null,
  "explanation_if_violating": null
}
```

Как использовать Structured Outputs с text.format
----------------------------------------------

Шаг 1: Определите вашу схему

Сначала вы должны разработать JSON-схему, которой должна следовать модель. Для справки см. [примеры](/docs/guides/structured-outputs#examples) в начале этого руководства.

Хотя Structured Outputs поддерживает большую часть JSON Schema, некоторые функции недоступны по соображениям производительности или техническим причинам. Для получения более подробной информации см. [здесь](/docs/guides/structured-outputs#supported-schemas).

#### Советы по составлению JSON-схемы

Чтобы максимизировать качество генераций модели, мы рекомендуем следующее:

*   Называйте ключи четко и интуитивно понятно.
*   Создавайте четкие заголовки и описания для важных ключей в вашей структуре.
*   Создавайте и используйте оценки (evals), чтобы определить структуру, которая лучше всего подходит для вашего случая использования.

Шаг 2: Укажите вашу схему в вызове API

Чтобы использовать Structured Outputs, просто укажите

```json
text: { format: { type: "json_schema", "strict": true, "schema": … } }
```

Например:

```python
response = client.responses.create(
    model="gpt-4o-2024-08-06",
    input=[
        {"role": "system", "content": "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом."},
        {"role": "user", "content": "как решить 8x + 7 = -23"}
    ],
    text={
        "format": {
            "type": "json_schema",
            "name": "math_response",
            "schema": {
                "type": "object",
                "properties": {
                    "steps": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "explanation": {"type": "string"},
                                "output": {"type": "string"}
                            },
                            "required": ["explanation", "output"],
                            "additionalProperties": False
                        }
                    },
                    "final_answer": {"type": "string"}
                },
                "required": ["steps", "final_answer"],
                "additionalProperties": False
            },
            "strict": True
        }
    }
)

print(response.output_text)
```

```javascript
const response = await openai.responses.create({
    model: "gpt-4o-2024-08-06",
    input: [
        { role: "system", content: "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом." },
        { role: "user", content: "как решить 8x + 7 = -23" }
    ],
    text: {
        format: {
            type: "json_schema",
            name: "math_response",
            schema: {
                type: "object",
                properties: {
                    steps: {
                        type: "array",
                        items: {
                            type: "object",
                            properties: {
                                explanation: { type: "string" },
                                output: { type: "string" }
                            },
                            required: ["explanation", "output"],
                            additionalProperties: false
                        }
                    },
                    final_answer: { type: "string" }
                },
                required: ["steps", "final_answer"],
                additionalProperties: false
            },
            strict: true
        }
    }
});

console.log(response.output_text);
```

```bash
curl https://api.openai.com/v1/responses \
  -H "Authorization: Bearer $OPENAI_API_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "model": "gpt-4o-2024-08-06",
    "input": [
      {
        "role": "system",
        "content": "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом."
      },
      {
        "role": "user",
        "content": "как решить 8x + 7 = -23"
      }
    ],
    "text": {
      "format": {
        "type": "json_schema",
        "name": "math_response",
        "schema": {
          "type": "object",
          "properties": {
            "steps": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "explanation": { "type": "string" },
                  "output": { "type": "string" }
                },
                "required": ["explanation", "output"],
                "additionalProperties": false
              }
            },
            "final_answer": { "type": "string" }
          },
          "required": ["steps", "final_answer"],
          "additionalProperties": false
        },
        "strict": true
      }
    }
  }'
```

**Примечание:** первый запрос с любой схемой будет иметь дополнительную задержку, поскольку наше API обрабатывает схему, но последующие запросы с той же схемой не будут иметь дополнительной задержки.

Шаг 3: Обработка крайних случаев

В некоторых случаях модель может не сгенерировать валидный ответ, соответствующий предоставленной JSON-схеме.

Это может произойти в случае отказа, если модель отказывается отвечать по соображениям безопасности, или если, например, достигнут лимит максимального количества токенов и ответ неполон.

```javascript
try {
  const response = await openai.responses.create({
    model: "gpt-4o-2024-08-06",
    input: [{
        role: "system",
        content: "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом.",
      },
      {
        role: "user",
        content: "как решить 8x + 7 = -23"
      },
    ],
    max_output_tokens: 50,
    text: {
      format: {
        type: "json_schema",
        name: "math_response",
        schema: {
          type: "object",
          properties: {
            steps: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  explanation: {
                    type: "string"
                  },
                  output: {
                    type: "string"
                  },
                },
                required: ["explanation", "output"],
                additionalProperties: false,
              },
            },
            final_answer: {
              type: "string"
            },
          },
          required: ["steps", "final_answer"],
          additionalProperties: false,
        },
        strict: true,
      },
    }
  });

  if (response.status === "incomplete" && response.incomplete_details.reason === "max_output_tokens") {
    // Обработайте случай, когда модель не вернула полный ответ
    throw new Error("Неполный ответ");
  }

  const math_response = response.output[0].content[0];

  if (math_response.type === "refusal") {
    // обработать отказ
    console.log(math_response.refusal);
  } else if (math_response.type === "output_text") {
    console.log(math_response.text);
  } else {
    throw new Error("Нет содержимого ответа");
  }
} catch (e) {
  // Обработайте крайние случаи
  console.error(e);
}
```

```python
try:
    response = client.responses.create(
        model="gpt-4o-2024-08-06",
        input=[
            {
                "role": "system",
                "content": "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом.",
            },
            {"role": "user", "content": "как решить 8x + 7 = -23"},
        ],
        text={
            "format": {
                "type": "json_schema",
                "name": "math_response",
                "strict": True,
                "schema": {
                    "type": "object",
                    "properties": {
                        "steps": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "explanation": {"type": "string"},
                                    "output": {"type": "string"},
                                },
                                "required": ["explanation", "output"],
                                "additionalProperties": False,
                            },
                        },
                        "final_answer": {"type": "string"},
                    },
                    "required": ["steps", "final_answer"],
                    "additionalProperties": False,
                },
                "strict": True,
            },
        },
    )
except Exception as e:
    # обработать ошибки, такие как finish_reason, refusal, content_filter и т.д.
    pass
```

### 

Отказы (Refusals) с Structured Outputs

При использовании Structured Outputs с пользовательским вводом модели OpenAI могут иногда отказываться выполнять запрос по соображениям безопасности. Поскольку отказ не обязательно следует схеме, которую вы указали в `response_format`, ответ API будет включать новое поле под названием `refusal`, указывающее на то, что модель отказалась выполнить запрос.

Когда свойство `refusal` появляется в вашем объекте вывода, вы можете отобразить отказ в вашем пользовательском интерфейсе или включить условную логику в код, который потребляет ответ, для обработки случая отказа.

```python
class Step(BaseModel):
    explanation: str
    output: str

class MathReasoning(BaseModel):
steps: list[Step]
final_answer: str

completion = client.chat.completions.parse(
model="gpt-4o-2024-08-06",
messages=[
{"role": "system", "content": "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом."},
{"role": "user", "content": "как решить 8x + 7 = -23"}
],
response_format=MathReasoning,
)

math_reasoning = completion.choices[0].message

# Если модель отказывается отвечать, вы получите сообщение об отказе

if (math_reasoning.refusal):
print(math_reasoning.refusal)
else:
print(math_reasoning.parsed)
```

```javascript
const Step = z.object({
explanation: z.string(),
output: z.string(),
});

const MathReasoning = z.object({
steps: z.array(Step),
final_answer: z.string(),
});

const completion = await openai.chat.completions.parse({
model: "gpt-4o-2024-08-06",
messages: [
{ role: "system", content: "Вы — полезный репетитор по математике. Проведите пользователя по решению шаг за шагом." },
{ role: "user", content: "как решить 8x + 7 = -23" },
],
response_format: zodResponseFormat(MathReasoning, "math_reasoning"),
});

const math_reasoning = completion.choices[0].message

// Если модель отказывается отвечать, вы получите сообщение об отказе
if (math_reasoning.refusal) {
console.log(math_reasoning.refusal);
} else {
console.log(math_reasoning.parsed);
}
```

Ответ API в случае отказа будет выглядеть примерно так:

```json
{
  "id": "resp_1234567890",
  "object": "response",
  "created_at": 1721596428,
  "status": "completed",
  "error": null,
  "incomplete_details": null,
  "input": [],
  "instructions": null,
  "max_output_tokens": null,
  "model": "gpt-4o-2024-08-06",
  "output": [{
    "id": "msg_1234567890",
    "type": "message",
    "role": "assistant",
    "content": [
      {
        "type": "refusal",
        "refusal": "Извините, я не могу помочь с этим запросом."
      }
    ]
  }],
  "usage": {
    "input_tokens": 81,
    "output_tokens": 11,
    "total_tokens": 92,
    "output_tokens_details": {
      "reasoning_tokens": 0,
    }
  },
}
```

### 

Советы и лучшие практики

#### Обработка пользовательского ввода

Если ваше приложение использует пользовательский ввод, убедитесь, что ваш промпт включает инструкции о том, как обрабатывать ситуации, когда ввод не может привести к валидному ответу.

Модель всегда будет пытаться соответствовать предоставленной схеме, что может привести к "галлюцинациям", если ввод полностью не связан со схемой.

Вы можете включить в промпт указание возвращать пустые параметры или определенную фразу, если модель обнаружит, что ввод несовместим с задачей.

#### Обработка ошибок

Structured Outputs все еще может содержать ошибки. Если вы видите ошибки, попробуйте скорректировать ваши инструкции, предоставить примеры в системных инструкциях или разбить задачи на более простые подзадачи. Обратитесь к [руководству по проектированию промптов (prompt engineering)](/docs/guides/prompt-engineering) для получения дополнительных указаний по настройке ваших входных данных.

#### Избегайте расхождения JSON-схемы и типов

Чтобы предотвратить расхождение между вашей JSON-схемой и соответствующими типами в вашем языке программирования, мы настоятельно рекомендуем использовать нативную поддержку Pydantic/Zod в SDK.

Если вы предпочитаете указывать JSON-схему напрямую, вы можете добавить правила CI, которые сигнализируют, когда редактируется либо JSON-схема, либо базовые объекты данных, или добавить шаг CI, который автоматически генерирует JSON-схему из определений типов (или наоборот).

Потоковая передача (Streaming)
---------

Вы можете использовать потоковую передачу для обработки ответов модели или аргументов вызова функций по мере их генерации и анализа их как структурированных данных.

Таким образом, вам не придется ждать завершения всего ответа, чтобы начать его обработку. Это особенно полезно, если вы хотите отображать поля JSON одно за другим или обрабатывать аргументы вызова функций, как только они становятся доступны.

Мы рекомендуем полагаться на SDK для обработки потоковой передачи с Structured Outputs.

```python
from typing import List

from openai import OpenAI
from pydantic import BaseModel

class EntitiesModel(BaseModel):
    attributes: List[str]
    colors: List[str]
    animals: List[str]

client = OpenAI()

with client.responses.stream(
    model="gpt-4.1",
    input=[
        {"role": "system", "content": "Извлеките сущности из входного текста"},
        {
            "role": "user",
            "content": "Быстрая коричневая лиса прыгает через ленивую собаку с пронзительными голубыми глазами",
        },
    ],
    text_format=EntitiesModel,
) as stream:
    for event in stream:
        if event.type == "response.refusal.delta":
            print(event.delta, end="")
        elif event.type == "response.output_text.delta":
            print(event.delta, end="")
        elif event.type == "response.error":
            print(event.error, end="")
        elif event.type == "response.completed":
            print("Завершено")
            # print(event.response.output)

    final_response = stream.get_final_response()
    print(final_response)
```

```javascript
import { OpenAI } from "openai";
import { zodTextFormat } from "openai/helpers/zod";
import { z } from "zod";

const EntitiesSchema = z.object({
  attributes: z.array(z.string()),
  colors: z.array(z.string()),
  animals: z.array(z.string()),
});

const openai = new OpenAI();
const stream = openai.responses
  .stream({
    model: "gpt-4.1",
    input: [
      { role: "user", content: "Какая сегодня погода в Париже?" },
    ],
    text: {
      format: zodTextFormat(EntitiesSchema, "entities"),
    },
  })
  .on("response.refusal.delta", (event) => {
    process.stdout.write(event.delta);
  })
  .on("response.output_text.delta", (event) => {
    process.stdout.write(event.delta);
  })
  .on("response.output_text.done", () => {
    process.stdout.write("\n");
  })
  .on("response.error", (event) => {
    console.error(event.error);
  });

const result = await stream.finalResponse();

console.log(result);
```

Поддерживаемые схемы
-----------------

Structured Outputs поддерживает подмножество языка [JSON Schema](https://json-schema.org/docs).

#### Поддерживаемые типы

Для Structured Outputs поддерживаются следующие типы:

*   String (Строка)
*   Number (Число)
*   Boolean (Логический)
*   Integer (Целое число)
*   Object (Объект)
*   Array (Массив)
*   Enum (Перечисление)
*   anyOf

#### Поддерживаемые свойства

В дополнение к указанию типа свойства вы можете указать выбор дополнительных ограничений:

**Поддерживаемые свойства `string`:**

*   `pattern` — Регулярное выражение, которому должна соответствовать строка.
*   `format` — Предопределенные форматы для строк. В настоящее время поддерживаются:
    *   `date-time`
    *   `time`
    *   `date`
    *   `duration`
    *   `email`
    *   `hostname`
    *   `ipv4`
    *   `ipv6`
    *   `uuid`

**Поддерживаемые свойства `number`:**

*   `multipleOf` — Число должно быть кратно этому значению.
*   `maximum` — Число должно быть меньше или равно этому значению.
*   `exclusiveMaximum` — Число должно быть меньше этого значения.
*   `minimum` — Число должно быть больше или равно этому значению.
*   `exclusiveMinimum` — Число должно быть больше этого значения.

**Поддерживаемые свойства `array`:**

*   `minItems` — Массив должен содержать как минимум указанное количество элементов.
*   `maxItems` — Массив должен содержать не более указанного количества элементов.

Вот несколько примеров использования этих ограничений типов:

Ограничения для строк

```json
{
    "name": "user_data",
    "strict": true,
    "schema": {
        "type": "object",
        "properties": {
            "name": {
                "type": "string",
                "description": "Имя пользователя"
            },
            "username": {
                "type": "string",
                "description": "Имя пользователя. Должно начинаться с @",
                "pattern": "^@[a-zA-Z0-9_]+$"
            },
            "email": {
                "type": "string",
                "description": "Email пользователя",
                "format": "email"
            }
        },
        "additionalProperties": false,
        "required": [
            "name", "username", "email"
        ]
    }
}
```

Ограничения для чисел

```json
{
    "name": "weather_data",
    "strict": true,
    "schema": {
        "type": "object",
        "properties": {
            "location": {
                "type": "string",
                "description": "Местоположение, для которого запрашивается погода"
            },
            "unit": {
                "type": ["string", "null"],
                "description": "Единица измерения температуры для возврата",
                "enum": ["F", "C"]
            },
            "value": {
                "type": "number",
                "description": "Фактическое значение температуры в указанном месте",
                "minimum": -130,
                "maximum": 130
            }
        },
        "additionalProperties": false,
        "required": [
            "location", "unit", "value"
        ]
    }
}
```

Обратите внимание, что эти ограничения [еще не поддерживаются для дообученных моделей (fine-tuned models)](/docs/guides/structured-outputs#some-type-specific-keywords-are-not-yet-supported).

#### Корневые объекты не должны быть `anyOf` и должны быть объектами

Обратите внимание, что корневой объект схемы должен быть объектом (object) и не должен использовать `anyOf`. Шаблон, который встречается в Zod (в качестве примера), — это использование размеченного объединения (discriminated union), которое создает `anyOf` на верхнем уровне. Поэтому код, подобный следующему, не будет работать:

```javascript
import { z } from 'zod';
import { zodResponseFormat } from 'openai/helpers/zod';

const BaseResponseSchema = z.object({/* ... */});
const UnsuccessfulResponseSchema = z.object({/* ... */});

const finalSchema = z.discriminatedUnion('status', [
BaseResponseSchema,
UnsuccessfulResponseSchema,
]);

// Недопустимая JSON-схема для Structured Outputs
const json = zodResponseFormat(finalSchema, 'final_schema');
```

#### Все поля должны быть `required`

Для использования Structured Outputs все поля или параметры функций должны быть указаны как `required` (обязательные).

```json
{
    "name": "get_weather",
    "description": "Получает погоду в заданном месте",
    "strict": true,
    "parameters": {
        "type": "object",
        "properties": {
            "location": {
                "type": "string",
                "description": "Местоположение, для которого запрашивается погода"
            },
            "unit": {
                "type": "string",
                "description": "Единица измерения температуры для возврата",
                "enum": ["F", "C"]
            }
        },
        "additionalProperties": false,
        "required": ["location", "unit"]
    }
}
```

Хотя все поля должны быть обязательными (и модель будет возвращать значение для каждого параметра), можно эмулировать необязательный параметр, используя объединенный тип (union type) с `null`.

```json
{
    "name": "get_weather",
    "description": "Получает погоду в заданном месте",
    "strict": true,
    "parameters": {
        "type": "object",
        "properties": {
            "location": {
                "type": "string",
                "description": "Местоположение, для которого запрашивается погода"
            },
            "unit": {
                "type": ["string", "null"],
                "description": "Единица измерения температуры для возврата",
                "enum": ["F", "C"]
            }
        },
        "additionalProperties": false,
        "required": [
            "location", "unit"
        ]
    }
}
```

#### Объекты имеют ограничения на глубину вложенности и размер

Схема может иметь до 5000 свойств объекта в общей сложности, с глубиной вложенности до 10 уровней.

#### Ограничения на общий размер строк

В схеме общая длина строк всех имен свойств, имен определений, значений перечислений и константных значений не может превышать 120 000 символов.

#### Ограничения на размер перечислений (enum)

Схема может содержать до 1000 значений перечислений во всех свойствах-перечислениях.

Для одного свойства-перечисления со строковыми значениями общая длина строк всех значений перечисления не может превышать 15 000 символов, когда значений перечисления больше 250.

#### `additionalProperties: false` должен всегда устанавливаться для объектов

`additionalProperties` определяет, допустимо ли для объекта содержать дополнительные ключи/значения, которые не были определены в JSON-схеме.

Structured Outputs поддерживает генерацию только указанных ключей/значений, поэтому мы требуем от разработчиков устанавливать `additionalProperties: false` для использования Structured Outputs.

```json
{
    "name": "get_weather",
    "description": "Получает погоду в заданном месте",
    "strict": true,
    "schema": {
        "type": "object",
        "properties": {
            "location": {
                "type": "string",
                "description": "Местоположение, для которого запрашивается погода"
            },
            "unit": {
                "type": "string",
                "description": "Единица измерения температуры для возврата",
                "enum": ["F", "C"]
            }
        },
        "additionalProperties": false,
        "required": [
            "location", "unit"
        ]
    }
}
```

#### Порядок ключей

При использовании Structured Outputs вывод будет производиться в том же порядке, в котором ключи расположены в схеме.

#### Некоторые ключевые слова, специфичные для типов, еще не поддерживаются

*   **Композиция:** `allOf`, `not`, `dependentRequired`, `dependentSchemas`, `if`, `then`, `else`

Для дообученных моделей мы дополнительно не поддерживаем следующее:

*   **Для строк:** `minLength`, `maxLength`, `pattern`, `format`
*   **Для чисел:** `minimum`, `maximum`, `multipleOf`
*   **Для объектов:** `patternProperties`
*   **Для массивов:** `minItems`, `maxItems`

Если вы включите Structured Outputs, указав `strict: true`, и вызовете API с неподдерживаемой JSON-схемой, вы получите ошибку.

#### Для `anyOf` вложенные схемы должны каждая быть валидной JSON-схемой в соответствии с этим подмножеством

Вот пример поддерживаемой схемы с `anyOf`:

```json
{
    "type": "object",
    "properties": {
        "item": {
            "anyOf": [
                {
                    "type": "object",
                    "description": "Объект пользователя для вставки в базу данных",
                    "properties": {
                        "name": {
                            "type": "string",
                            "description": "Имя пользователя"
                        },
                        "age": {
                            "type": "number",
                            "description": "Возраст пользователя"
                        }
                    },
                    "additionalProperties": false,
                    "required": [
                        "name",
                        "age"
                    ]
                },
                {
                    "type": "object",
                    "description": "Объект адреса для вставки в базу данных",
                    "properties": {
                        "number": {
                            "type": "string",
                            "description": "Номер адреса. Например, для '123 main st' это будет '123'"
                        },
                        "street": {
                            "type": "string",
                            "description": "Название улицы. Например, для '123 main st' это будет 'main st'"
                        },
                        "city": {
                            "type": "string",
                            "description": "Город адреса"
                        }
                    },
                    "additionalProperties": false,
                    "required": [
                        "number",
                        "street",
                        "city"
                    ]
                }
            ]
        }
    },
    "additionalProperties": false,
    "required": [
        "item"
    ]
}
```

#### Определения (Definitions) поддерживаются

Вы можете использовать определения (`$defs`) для определения подсхем, на которые ссылаются throughout вашей схемы. Ниже приведен простой пример.

```json
{
    "type": "object",
    "properties": {
        "steps": {
            "type": "array",
            "items": {
                "$ref": "#/$defs/step"
            }
        },
        "final_answer": {
            "type": "string"
        }
    },
    "$defs": {
        "step": {
            "type": "object",
            "properties": {
                "explanation": {
                    "type": "string"
                },
                "output": {
                    "type": "string"
                }
            },
            "required": [
                "explanation",
                "output"
            ],
            "additionalProperties": false
        }
    },
    "required": [
        "steps",
        "final_answer"
    ],
    "additionalProperties": false
}
```

#### Рекурсивные схемы поддерживаются

Пример рекурсивной схемы с использованием `#` для указания корневой рекурсии.

```json
{
    "name": "ui",
    "description": "Динамически сгенерированный пользовательский интерфейс",
    "strict": true,
    "schema": {
        "type": "object",
        "properties": {
            "type": {
                "type": "string",
                "description": "Тип компонента UI",
                "enum": ["div", "button", "header", "section", "field", "form"]
            },
            "label": {
                "type": "string",
                "description": "Метка компонента UI, используется для кнопок или полей формы"
            },
            "children": {
                "type": "array",
                "description": "Вложенные компоненты UI",
                "items": {
                    "$ref": "#"
                }
            },
            "attributes": {
                "type": "array",
                "description": "Произвольные атрибуты для компонента UI, подходящие для любого элемента",
                "items": {
                    "type": "object",
                    "properties": {
                        "name": {
                            "type": "string",
                            "description": "Имя атрибута, например onClick или className"
                        },
                        "value": {
                            "type": "string",
                            "description": "Значение атрибута"
                        }
                    },
                    "additionalProperties": false,
                    "required": ["name", "value"]
                }
            }
        },
        "required": ["type", "label", "children", "attributes"],
        "additionalProperties": false
    }
}
```

Пример рекурсивной схемы с явной рекурсией:

```json
{
    "type": "object",
    "properties": {
        "linked_list": {
            "$ref": "#/$defs/linked_list_node"
        }
    },
    "$defs": {
        "linked_list_node": {
            "type": "object",
            "properties": {
                "value": {
                    "type": "number"
                },
                "next": {
                    "anyOf": [
                        {
                            "$ref": "#/$defs/linked_list_node"
                        },
                        {
                            "type": "null"
                        }
                    ]
                }
            },
            "additionalProperties": false,
            "required": [
                "next",
                "value"
            ]
        }
    },
    "additionalProperties": false,
    "required": [
        "linked_list"
    ]
}
```

JSON mode
---------

JSON mode — это более базовая версия функции Structured Outputs. В то время как JSON mode гарантирует, что вывод модели является валидным JSON, Structured Outputs надежно сопоставляет вывод модели с указанной вами схемой. Мы рекомендуем использовать Structured Outputs, если это поддерживается для вашего случая использования.

Когда JSON mode включен, вывод модели гарантированно является валидным JSON, за исключением некоторых крайних случаев, которые вы должны обнаруживать и обрабатывать соответствующим образом.

Чтобы включить JSON mode в Responses API, вы можете установить `text.format` в `{ "type": "json_object" }`. Если вы используете вызов функций, JSON mode всегда включен.

Важные замечания:

*   При использовании JSON mode вы всегда должны инструктировать модель генерировать JSON через какое-либо сообщение в разговоре, например, через ваше системное сообщение. Если вы не включите явную инструкцию генерировать JSON, модель может генерировать бесконечный поток пробелов, и запрос может выполняться непрерывно, пока не достигнет лимита токенов. Чтобы помочь избежать этого, API выдаст ошибку, если строка "JSON" не появится где-нибудь в контексте.
*   JSON mode не гарантирует, что вывод соответствует какой-либо конкретной схеме, а только то, что он валиден и парсится без ошибок. Вам следует использовать Structured Outputs для гарантии соответствия вашей схеме, или, если это невозможно, использовать библиотеку валидации и, возможно, повторные попытки, чтобы обеспечить соответствие вывода желаемой схеме.
*   Ваше приложение должно обнаруживать и обрабатывать крайние случаи, которые могут привести к тому, что вывод модели не будет полным JSON-объектом (см. ниже)

Обработка крайних случаев

```javascript
const we_did_not_specify_stop_tokens = true;

try {
  const response = await openai.responses.create({
    model: "gpt-3.5-turbo-0125",
    input: [
      {
        role: "system",
        content: "Вы — полезный ассистент, предназначенный для вывода JSON.",
      },
      { role: "user", content: "Кто выиграл мировую серию в 2020 году? Пожалуйста, ответьте в формате {winner: ...}" },
    ],
    text: { format: { type: "json_object" } },
  });

  // Проверить, не была ли беседа слишком длинной для контекстного окна, что привело к неполному JSON
  if (response.status === "incomplete" && response.incomplete_details.reason === "max_output_tokens") {
    // ваш код должен обработать этот случай ошибки
  }

  // Проверить, не отказала ли система безопасности OpenAI в запросе и не сгенерировала ли отказ вместо ответа
  if (response.output[0].content[0].type === "refusal") {
    // ваш код должен обработать этот случай ошибки
    // В этом случае поле .content будет содержать объяснение (если есть), которое модель сгенерировала для причины отказа
    console.log(response.output[0].content[0].refusal)
  }

  // Проверить, не содержал ли вывод модели ограниченный контент, из-за чего генерация JSON была прервана и может быть частичной
  if (response.status === "incomplete" && response.incomplete_details.reason === "content_filter") {
    // ваш код должен обработать этот случай ошибки
  }

  if (response.status === "completed") {
    // В этом случае модель либо успешно завершила генерацию JSON-объекта в соответствии с вашей схемой, либо модель сгенерировала один из токенов, которые вы предоставили как "стоп-токен"

    if (we_did_not_specify_stop_tokens) {
      // Если вы не указали никаких стоп-токенов, то генерация завершена, и ключ content будет содержать сериализованный JSON-объект
      // Это успешно распарсится и теперь должно содержать {"winner": "Los Angeles Dodgers"}
      console.log(JSON.parse(response.output_text))
    } else {
      // Проверить, заканчивается ли response.output_text одним из ваших стоп-токенов, и обработать соответствующим образом
    }
  }
} catch (e) {
  // Ваш код должен обрабатывать ошибки здесь, например, сетевую ошибку при вызове API
  console.error(e)
}
```

```python
we_did_not_specify_stop_tokens = True

try:
    response = client.responses.create(
        model="gpt-3.5-turbo-0125",
        input=[
            {"role": "system", "content": "Вы — полезный ассистент, предназначенный для вывода JSON."},
            {"role": "user", "content": "Кто выиграл мировую серию в 2020 году? Пожалуйста, ответьте в формате {winner: ...}"}
        ],
        text={"format": {"type": "json_object"}}
    )

    # Проверить, не была ли беседа слишком длинной для контекстного окна, что привело к неполному JSON
    if response.status == "incomplete" and response.incomplete_details.reason == "max_output_tokens":
        # ваш код должен обработать этот случай ошибки
        pass

    # Проверить, не отказала ли система безопасности OpenAI в запросе и не сгенерировала ли отказ вместо ответа
    if response.output[0].content[0].type == "refusal":
        # ваш код должен обработать этот случай ошибки
        # В этом случае поле .content будет содержать объяснение (если есть), которое модель сгенерировала для причины отказа
        print(response.output[0].content[0]["refusal"])

    # Проверить, не содержал ли вывод модели ограниченный контент, из-за чего генерация JSON была прервана и может быть частичной
    if response.status == "incomplete" and response.incomplete_details.reason == "content_filter":
        # ваш код должен обработать этот случай ошибки
        pass

    if response.status == "completed":
        # В этом случае модель либо успешно завершила генерацию JSON-объекта в соответствии с вашей схемой, либо модель сгенерировала один из токенов, которые вы предоставили как "стоп-токен"

        if we_did_not_specify_stop_tokens:
            # Если вы не указали никаких стоп-токенов, то генерация завершена, и ключ content будет содержать сериализованный JSON-объект
            # Это успешно распарсится и теперь должно содержать "{"winner": "Los Angeles Dodgers"}"
            print(response.output_text)
        else:
            # Проверить, заканчивается ли response.output_text одним из ваших стоп-токенов, и обработать соответствующим образом
            pass
except Exception as e:
    # Ваш код должен обрабатывать ошибки здесь, например, сетевую ошибку при вызове API
    print(e)
```

Ресурсы
---------

Чтобы узнать больше о Structured Outputs, мы рекомендуем изучить следующие ресурсы:

*   Ознакомьтесь с нашим [вводным кулинарной книгой (cookbook)](https://cookbook.openai.com/examples/structured_outputs_intro) по Structured Outputs
*   Узнайте, [как создавать мульти-агентные системы](https://cookbook.openai.com/examples/structured_outputs_multi_agent) с помощью Structured Outputs
