---
id: 9f147d3a-4cb5-4c6f-a635-ed7f8046502d
name: Chain of Thought Prompting
title: "Chain of Thought Prompting: Guiding AI to Think Step by Step"
description: Руководство по CoT prompting и пошаговому рассуждению
tags:
  - prompt-engineering
  - step-by-step
  - thinking
  - cot
  - chain-of-throught
  - third-party-docs
authors:
  - zilliz
url: https://zilliz.com/glossary/chain-of-thoughts
language: en
translated: false
---
# Chain of Thought Prompting: Guiding AI to Think Step by Step

## Оглавление

- [[#What is Chain-of-Thought Prompting?|What is Chain-of-Thought Prompting?]]
- [[#How Chain-of-Thought Prompting Works?|How Chain-of-Thought Prompting Works?]]
	- [[#How Chain-of-Thought Prompting Works?#Practical Example: Before and After Chain-of-Thought Prompting|Practical Example: Before and After Chain-of-Thought Prompting]]
- [[#Chain-of-Thought Variants with Prompt Examples|Chain-of-Thought Variants with Prompt Examples]]
	- [[#Chain-of-Thought Variants with Prompt Examples#1. Standard Chain-of-Thought Prompting|1. Standard Chain-of-Thought Prompting]]
	- [[#Chain-of-Thought Variants with Prompt Examples#2. Few-Shot Chain-of-Thought Prompting|2. Few-Shot Chain-of-Thought Prompting]]
	- [[#Chain-of-Thought Variants with Prompt Examples#3. Self-Consistency Chain-of-Thought|3. Self-Consistency Chain-of-Thought]]
	- [[#Chain-of-Thought Variants with Prompt Examples#4. Interactive Chain-of-Thought|4. Interactive Chain-of-Thought]]
	- [[#Chain-of-Thought Variants with Prompt Examples#5. Multi-Agent Chain-of-Thought|5. Multi-Agent Chain-of-Thought]]
	- [[#Chain-of-Thought Variants with Prompt Examples#6. Recursive Chain-of-Thought|6. Recursive Chain-of-Thought]]
- [[#When to Use Chain-of-Thought Prompting|When to Use Chain-of-Thought Prompting]]
- [[#Further Improving Chain-of-Thought Prompting|Further Improving Chain-of-Thought Prompting]]
- [[#Limitations of Chain-of-Thought Prompting|Limitations of Chain-of-Thought Prompting]]
- [[#Real-World Use Cases of Chain-of-Thought Prompting|Real-World Use Cases of Chain-of-Thought Prompting]]
	- [[#Real-World Use Cases of Chain-of-Thought Prompting#1. Mathematical Problem Solving|1. Mathematical Problem Solving]]
	- [[#Real-World Use Cases of Chain-of-Thought Prompting#2. Legal Reasoning and Contract Analysis|2. Legal Reasoning and Contract Analysis]]
	- [[#Real-World Use Cases of Chain-of-Thought Prompting#3. Customer Support and Troubleshooting|3. Customer Support and Troubleshooting]]
	- [[#Real-World Use Cases of Chain-of-Thought Prompting#4. Medical Diagnosis and Decision Support|4. Medical Diagnosis and Decision Support]]
	- [[#Real-World Use Cases of Chain-of-Thought Prompting#5. Complex Question Answering|5. Complex Question Answering]]
	- [[#Real-World Use Cases of Chain-of-Thought Prompting#6. Game Strategy and Puzzle Solving|6. Game Strategy and Puzzle Solving]]
- [[#How Can Milvus Boost the Efficiency of Chain-of-Thought Prompting?|How Can Milvus Boost the Efficiency of Chain-of-Thought Prompting?]]
- [[#Conclusion|Conclusion]]
- [[#FAQs on Chain-of-Thought Prompting|FAQs on Chain-of-Thought Prompting]]
- [[#Related Resources|Related Resources]]


## What is Chain-of-Thought Prompting?

Chain-of-thought (CoT) is a [prompt engineering](https://zilliz.com/glossary/prompt-as-code-\(prompt-engineering\)) technique that helps large [Large Language Models (LLMs)](https://zilliz.com/glossary/large-language-models-\(llms\)) break down complex problems into smaller, easier steps. Instead of giving a direct answer, the model walks through a sequence of thoughts or reasoning, like how a person might solve a problem step by step.

## How Chain-of-Thought Prompting Works?

Chain-of-thought prompting works by guiding [AI models](https://zilliz.com/glossary/foundation-models) to tackle complex problems step-by-step. This method is based on sequential reasoning; each step builds on the previous one. For example, suppose the model is asked a difficult math problem. Instead of trying to give the final answer right away, it first works through smaller calculations or logical steps, leading to a more accurate final result. This process of breaking down the problem helps the model get the right answer and understand how it arrived there, improving the overall response quality.

### Practical Example: Before and After Chain-of-Thought Prompting

Let’s understand how CoT works through an example in action. The two prompts show how chain-of-thought prompting impacts the AI model’s response.

**1. Prompt Before CoT**

_Categorize the following user reviews as 'Positive,' 'Negative,' or 'Neutral':_

_"The product arrived late but works fine."_

_"Absolutely fantastic service and quality!"_

_"I didn't like the color, but customer support was helpful."_

_"Terrible experience. Would not recommend."_

_"It's okay, does the job.""_

**Response:**

![Prompt without COT.png](https://assets.zilliz.com/answer_with_prompt_before_COT_d2ba394661.png)Prompt without COT.png

Figure: Prompt without CoT

**Analysis:**

In the naive prompt, the model provides a basic output where each review is simply categorized without explaining how it arrived at each decision. It doesn’t show the user the reasoning behind the categorization or provide insight into which specific words or phrases were considered positive or negative. The output is correct but lacks depth and transparency.

**2. Prompt After CoT**

- _Please categorize the following user reviews as 'Positive,' 'Negative,' or 'Neutral.' For each review, follow these steps:_

_Read the review carefully._

1. _Identify any positive words or phrases._
    
2. _Identify any negative words or phrases._
    
3. _Assess the overall sentiment based on the positive and negative elements._
    
4. _Assign the appropriate category._
    
5. _In output, present all the information in a table._
    

_Proceed with each review:_

_"The product arrived late but works fine."_

_"Absolutely fantastic service and quality!"_

_"I didn't like the color, but customer support was helpful."_

_"Terrible experience. Would not recommend."_

_"It's okay, does the job."_

**Response:**

![Prompt with CoT](https://assets.zilliz.com/answer_with_COT_d295b92fd4.png)Prompt with CoT

Figure: Prompt with CoT

**Analysis:**

- The model explains its reasoning by identifying positive and negative phrases in each review, giving the user a clear understanding of the decision.
    
- A table is generated, which not only lists the final category but also shows the key sentiment phrases for each review, making the output more structured and transparent.
    

## Chain-of-Thought Variants with Prompt Examples

Here are different types of Chain-of-Thought prompting with practical, prompt examples for each:

### 1. Standard Chain-of-Thought Prompting

This type of chain of thoughts asks the model to think through each part of the problem step by step.

**Prompt Example:**

_A train moving at 60 miles per hour needs to travel 180 miles. How long will it take to complete the journey? Let’s go through it step by step._

**Expected Model Response:**

1. "The train is traveling at 60 miles per hour."
    
2. "The distance to cover is 180 miles."
    
3. "To find the time, we need to divide the distance by the speed: 180 ÷ 60 = 3."
    
4. "So, it will take 3 hours to reach the destination."
    

### 2. Few-Shot Chain-of-Thought Prompting

In a few-shot CoT, the model is given a few examples before being asked to solve a new problem.

**Prompt Example:**

_"Here’s an example: What is its speed if a car covers 100 miles in 2 hours? The speed is 100 ÷ 2 = 50 miles per hour.  
Here’s another example: If a plane covers 600 miles in 3 hours, the speed is 600 ÷ 3 = 200 miles per hour.  
Now, using the same steps, if a cyclist covers 90 miles in 3 hours, what is their speed?"_

**Expected Model Response:**

1. "The distance is 90 miles, and the time is 3 hours."
    
2. "To find the speed, we divide 90 by 3."
    
3. "90 ÷ 3 = 30 miles per hour."
    

### 3. Self-Consistency Chain-of-Thought

The model generates multiple reasoning paths and picks the most consistent answer.

**Prompt Example:**

_"If a person buys 6 apples for $12, how much does each apple cost? Let’s try solving this problem in different ways and compare the answers."_ **Expected Model Response:**

Path 1:

1. "The total cost is $12, and there are 6 apples."
2. "$12 ÷ 6 = $2 per apple." Path 2:
3. "If 6 apples cost $12, then each apple costs $2 because $12 divided by 6 equals $2." **Final Answer:** "Both methods lead to the same result, so the cost per apple is $2."

### 4. Interactive Chain-of-Thought

The model asks for clarification or extra information when needed in this variant. **Prompt Example:** _"You have a total of 50 dollars and want to buy some books. Each book costs 7 dollars. How many books can you buy, and how much money will be left over? Please ask if you need more information before solving."_ **Expected Model Response:** 1. "Are there taxes included in the book price, or is it just $7 per book?" (If the user confirms it's $7 per book without taxes)  
2. "Alright, without taxes, $50 divided by $7 gives roughly 7 books. Since 7 × 7 equals 49." 3. "You’ll have $1 remaining after purchasing 7 books."

### 5. Multi-Agent Chain-of-Thought

In this case, multiple AI agents work together to solve the problem by tackling different parts. **Prompt Example:** _"There are 45 students and 9 teachers going on a trip. If each bus can hold 15 people, how many buses are needed? Let's have one agent calculate the number of total people, another agent calculate the division, and a third one verify the solution."_ **Expected Model Response:** **Agent 1:** "The total number of people is 45 students + 9 teachers = 54 people."  
**Agent 2:** "We divide 54 people by 15 per bus: 54 ÷ 15 = 3.6, which means we need 4 buses."  
**Agent 3:** “The total number of people and the calculation are accurate." We need 4 buses."

### 6. Recursive Chain-of-Thought

The model checks its own reasoning and revises steps if necessary. **Prompt Example:** **“**_With \$120 in hand, you spend $45 on groceries, $30 on gas, and $20 on entertainment. How much money do you have left? Please check your calculations as you go."_ **Expected Model Response:**

1. "Starting with $120."
2. "$120 - $45 = $75."
3. "$75 - $30 = $45."
4. "$45 - $20 = $25."
5. "Checking: $120 - $45 = $75, then $75 - $30 = $45, and finally $45 - $20 = $25. The answer is $25."

## When to Use Chain-of-Thought Prompting

Chain-of-thought prompting is suitable for tasks that require multi-step reasoning, logical thinking, or careful analysis. Below are some scenarios in which CoT can be useful:

1. **Math Problems or Calculations:** When the task involves multiple operations or calculations, CoT guides the model through each step. For example, solving word problems or complex arithmetic requires breaking down the information into smaller, manageable pieces before performing calculations.
    
2. **Logical Reasoning and Deduction:** CoT is great for tasks where the model needs to analyze information step by step, whether making a decision based on several factors or solving a puzzle involving logical steps.
    
3. **Multi-Step Problem Solving:** Tasks like balancing a budget, following a recipe, or solving a technical issue often require several steps. CoT helps the model follow these steps in the right order.
    
4. **Complex Question Answering:** CoT can guide the model through the reasoning required to give a more accurate and detailed response when asked complex questions in science, law, or philosophy. Instead of guessing, it analyzes the facts and logic needed to form a sound answer.
    

## Further Improving Chain-of-Thought Prompting

To further enhance the effectiveness of the prompt, some other techniques can be combined to handle reasoning. For example:

1. **Providing Clear, Structured Prompts:** The quality of the prompt's quality significantly impacts reasoning. Prompts should be designed to enhance CoT to break problems into logical steps. The more structured and detailed the prompt, the better the model will follow the reasoning process.
    
2. **Combining CoT with Few-Shot Learning:** Few-shot learning, where the model gives a few examples of reasoning through problems, can improve CoT. By showing the model several similar cases where step-by-step reasoning was used, it can better understand how to approach new tasks.
    
3. **Self-Reflection and Verification:** One way to improve CoT is to prompt the model to check its own reasoning. After generating an answer, the model can be asked to review its steps to see if they make sense. This helps catch any logical mistakes or gaps before giving the final answer.
    
4. **Incorporating Self-Consistency:** Self-consistency is a method where the model generates multiple reasoning paths and then compares them to see if they lead to the same conclusion. If multiple paths agree, the model is more likely to have arrived at the correct answer. For example, the model can be asked to solve a problem in two different ways, and then the most consistent answer can be selected across different approaches.
    

## Limitations of Chain-of-Thought Prompting

While Chain-of-Thought prompting is a powerful approach for improving AI reasoning, it does have some limitations that can impact its effectiveness in certain scenarios.

**1. Susceptibility to Errors in Reasoning:** CoT depends on the model following a logical, step-by-step process. However, if the model makes a mistake in one step, that error can carry through to the final answer.

**2. Overhead in Time and Resources:** CoT requires more computational resources because the model has to think through each step instead of providing a direct answer. This can make the process slower and more resource-intensive, especially for complex or multi-step problems. In situations where speed is a priority, CoT might not be ideal.

**4. Dependence on Well-Designed Prompts:** The effectiveness of CoT heavily depends on the quality of the prompts provided. If the prompt isn't clear or well-structured, the model might struggle to reason through the problem. Crafting these prompts requires effort and expertise.

**5. Limited Generalization to Unfamiliar Tasks:** CoT is highly effective for tasks that it has seen before or that closely resemble previous tasks. However, when presented with unfamiliar problems or tasks outside its training data, the model may struggle to apply CoT effectively, as it relies on learned reasoning patterns.

**6. Risk of Overfitting to Prompt Structure:** Over time, a model trained to use CoT could become overly reliant on specific prompts, limiting flexibility. [Overfitting](https://zilliz.com/learn/understanding-regularization-in-nueral-networks) might expect problems to always be presented in a particular format, making it harder for the model to adapt to new or differently phrased tasks.

## Real-World Use Cases of Chain-of-Thought Prompting

Chain-of-thought prompting has a wide range of practical applications across various fields. Here are some key real-world use cases where CoT can be highly beneficial:

### 1. Mathematical Problem Solving

CoT is extremely useful in mathematics education and tutoring platforms. Students can understand the process by breaking down math problems into smaller, logical steps instead of just getting the final answer. It is also helpful for advanced calculations in areas like algebra, calculus, and statistics.

### 2. Legal Reasoning and Contract Analysis

In legal systems, CoT helps AI systems evaluate legal documents, analyze clauses, and generate legal advice by systematically going through each point in a contract or case. It can also explain step-by-step [legal reasoning](https://zilliz.com/blog/simplifying-legal-research-with-rag-milvus-ollama) that makes AI more transparent and trustworthy in legal processes.

### 3. Customer Support and Troubleshooting

CoT helps AI-driven chatbots or support systems guide users through step-by-step troubleshooting processes. This is useful for technical issues, where the user must follow instructions or diagnostic steps to resolve a problem.

### 4. Medical Diagnosis and Decision Support

CoT can help doctors or medical professionals analyze symptoms, test results, and medical history to suggest possible diagnoses or treatments, explaining how each conclusion is reached.

### 5. Complex Question Answering

CoT is highly effective in answering complex questions that require more than just retrieving a fact. For questions in fields like history, science, or law, CoT can help AI systems provide detailed, multi-step answers explaining the response's reasoning.

### 6. Game Strategy and Puzzle Solving

In gaming, strategies are generated by considering multiple steps in a sequence. CoT breaks down complex strategies in gaming or puzzle-solving scenarios into smaller, thoughtful moves that lead to better decision-making and gameplay for games like chess, Go, or puzzle-based games.

## How Can Milvus Boost the Efficiency of Chain-of-Thought Prompting?

[Milvus](https://zilliz.com/what-is-milvus), an open-source [vector database](https://zilliz.com/learn/what-is-vector-database) developed by [Zilliz](https://zilliz.com/cloud), is designed to efficiently store and retrieve [unstructured data](https://zilliz.com/learn/introduction-to-unstructured-data) like images, text, and video. While Chain-of-Thought prompting focuses on improving AI models' reasoning abilities, Milvus enhances how these models manage and process large-scale vector data.

1. **Efficient Data Retrieval for Complex Reasoning:** CoT relies on AI models having access to relevant information for step-by-step reasoning. Milvus is an efficient backend, storing vast amounts of vector data (like text embeddings) and providing quick retrieval. This allows AI models to access the data they need at each step of the reasoning process without delays.

2. **Supporting Large-Scale Applications:** CoT prompts require handling extensive datasets for multi-step reasoning in many real-world cases. Milvus’s s in many real-world cases enables AI models to work with large datasets without compromising on speed or performance.
    
3. **Optimized Similarity Searches:** Milvus is built for fast [semantic searches](https://zilliz.com/glossary/semantic-search) and [similarity searches](https://zilliz.com/learn/vector-similarity-search), which enhances CoT by allowing AI to access semantically related data quickly. This speeds up the reasoning process, as the model can pull relevant information more accurately and efficiently when solving multi-step problems.
    

## Conclusion

In summary, Chain-of-Thought Prompting helps AI models tackle complex problems by breaking them into logical steps, improving accuracy and clarity. Milvus enhances this process by enabling quick access to large amounts of unstructured data so that the AI can pull relevant information as it works through each step. CoT and Milvus offer practical solutions for handling complex tasks across fields like research, customer support, and financial analysis, making AI more effective and reliable in real-world applications.

## FAQs on Chain-of-Thought Prompting

- ### **How does Chain-of-Thought Prompting improve AI reasoning?**
    

Chain-of-thought prompting improves AI reasoning by guiding the model through problems step by step. This method encourages the model to break down complex tasks into smaller, manageable pieces, reducing errors and enhancing accuracy.

- ### **When should Chain-of-Thought Prompting be used?**
    

CoT is best used for tasks that require multi-step reasoning, deep logical analysis, or complex problem-solving. Examples include math problems, logical deductions, technical troubleshooting, and multi-faceted decision-making processes.

- ### **What are the key benefits of Chain-of-Thought Prompting?**
    

The key benefits of CoT include improved accuracy, better handling of complex problems, reduced errors, enhanced model transparency, and a structured approach that makes AI reasoning more understandable and reliable.

- ### **How does Milvus improve the efficiency of Chain-of-Thought Prompting?**
    

Milvus enhances CoT prompting by efficiently storing and retrieving large-scale unstructured data, such as text and images. It enables AI models to access relevant data quickly at each reasoning stage for smooth and fast performance in complex, multi-step tasks.

- ### **How does Chain-of-Thought Prompting differ from traditional AI responses?**
    

Traditional AI responses often attempt to provide an answer directly without detailing the reasoning process. On the other hand, Chain-of-Thought Prompting guides the model in explaining its reasoning step by step, offers transparency, and follows a logical progression toward the solution.

## Related Resources

- [[2201.11903] Chain-of-Thought Prompting Elicits Reasoning in Large Language Models](https://arxiv.org/abs/2201.11903)
    
- [Introduction to Prompt Engineering](https://zilliz.com/glossary/prompt-as-code-\(prompt-engineering\))
    
- [An LLM Powered Text to Image Prompt Generation with Milvus](https://zilliz.com/blog/llm-powered-text-to-image-prompt-generation-with-milvus)
    
- [Prompting in LangChain](https://zilliz.com/blog/prompting-langchain)
    
- [Prompt Chaining](https://zilliz.com/glossary/prompt-chaining)
    
- [Everything You Need to Know About Zero Shot Learning](https://zilliz.com/learn/what-is-zero-shot-learning)
    
- [ChatGPT+ Vector database + prompt-as-code - The CVP Stack](https://zilliz.com/blog/ChatGPT-VectorDB-Prompt-as-code)
    
- [Exploring LLM-Driven Agents in the Age of AI](https://zilliz.com/blog/explore-llm-driven-agents-in-age-of-AI)