---
title: How to write better prompts for GitHub Copilot
description: In this prompt guide for GitHub Copilot, two GitHub developer advocates, Rizel and Michelle, will share examples and best practices for communicating your desired results to the AI pair programmer.
domain: prompt-engineering
type: tutorial
tags:
  - prompt-engineering
  - copilot-github
  - llm
author:
  - Rizel Scarlett
  - Michelle Duke
url: https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text
language: EN
---


Get the latest on GitHub Copilot

GitHub Copilot is always adding new features and functionality. Check out these resources to keep up to date:

- Our [comprehensive guide on how to use GitHub Copilot with real-world examples](https://github.blog/developer-skills/github/what-can-github-copilot-do-examples/)
- The [GitHub Copilot tag](https://github.blog/developer-skills/github/what-can-github-copilot-do-examples/) and [GitHub Copilot category](https://github.blog/ai-and-ml/github-copilot/) on the GitHub Blog
- The [GitHub Copilot documentation](https://docs.github.com/en/copilot) and the [GitHub Copilot Chat cookbook](https://docs.github.com/en/copilot/example-prompts-for-github-copilot-chat)

Here’s an example of GitHub Copilot generating an irrelevant solution:

![When we wrote this prompt to GitHub Copilot, 'draw an ice cream cone with ice cream using p5.js,' the AI pair programmer generated an image that looked like a bulls-eye target sitting on top of a stand.](https://github.blog/wp-content/uploads/2023/06/245577016-2c22e7a6-6d99-47e2-8f45-1322ac2d2e1b.gif)

When we adjusted our prompt, we were able to generate more accurate results:

![When we wrote this prompt to GitHub Copilot, 'draw an ice cream cone with an ice cream scoop and a cherry on top,' and specified details about each part of the picture, GitHub Copilot generated a picture of the ice cream cone that we wanted. Here were the prompts for those details: 1) 'The ice cream cone will be a triangle with the point facing down, wider point at the top. It should have light brown fill' 2) 'The ice cream scoop will be a half circle on top of the cone with a light pink fill' 3) 'The cherry will be a circle on top of the ice cream scoop with a red fill' 4) 'Light blue background'](https://github.blog/wp-content/uploads/2023/06/245577026-e97a29d4-1378-48c1-a8ff-6adbc3bb2d8d.gif)

We’re both developers and AI enthusiasts ourselves. I, [Rizel](https://github.com/blackgirlbytes), have used GitHub Copilot to build a [browser extension](https://github.blog/2023-05-12-how-i-used-github-copilot-to-build-a-browser-extension/), [rock, paper, scissors game](https://github.com/blackgirlbytes/rock-paper-scissors-copilot), and to [send a Tweet](https://dev.to/github/how-to-send-a-tweet-with-github-copilot-4ih7). And I, [Michelle](https://github.com/mishmanners), launched an AI company in 2016. We’re both developer advocates at GitHub and love to share our top tips for working with GitHub Copilot.

In this guide for GitHub Copilot, we’ll cover:

- [What exactly a prompt is and what prompt engineering is, too](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#whats-a-prompt-and-what-is-prompt-engineering) (hint: it depends on whether you’re talking to a developer or a machine learning researcher).
- [Three best practices](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#3-best-practices-for-prompt-crafting-with-github-copilot) and [three additional tips for prompt crafting with GitHub Copilot](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#three-additional-tips-for-prompt-crafting-with-github-copilot).
- [An example](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#a-practice-example) where you can try your hand at prompting GitHub Copilot to assist you in building a browser extension.

### [Progress over perfection](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#progress-over-perfection)

Even with our experience using AI, we recognize that everyone is in a trial and error phase with generative AI technology. We also know the challenge of providing generalized prompt-crafting tips because models vary, as do the individual problems that developers are working on. This isn’t an end-all, be-all guide. Instead, we’re sharing what we’ve learned about prompt crafting to accelerate collective learning during this new age of software development.

## [What’s a prompt and what is prompt engineering?](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#whats-a-prompt-and-what-is-prompt-engineering)

It depends on who you talk to.

In the context of generative AI coding tools, **a prompt can mean different things, depending on whether you’re asking machine learning (ML) researchers** who are building and fine-tuning these tools, **or developers** who are using them in their IDEs.

**For this guide, we’ll define the terms from the point of view of a developer who’s using a generative AI coding tool in the IDE**. But to give you the full picture, we also added the ML researcher definitions below in our chart.

| **Prompts**   | **Prompt engineering**                                                                                                                                                                          | **Context**                                                                                                                                                                                                                                                  |                                                                                                                                                                                                                                                                                                                                                |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Developer     | **Code blocks, individual lines of code, or natural language comments a developer writes** to generate a specific suggestion from GitHub Copilot                                                | **Providing instructions or comments in the IDE** to generate specific coding suggestions                                                                                                                                                                    | **Details that are provided by a developer** to specify the desired output from a generative AI coding tool                                                                                                                                                                                                                                    |
| ML researcher | **Compilation of IDE code and relevant context** (IDE comments, code in open files, etc.) **that is continuously generated by algorithms** and sent to the model of a generative AI coding tool | [Creating algorithms that will generate prompts (compilations of IDE code and context) for a large language model](https://github.blog/2023-05-17-how-github-copilot-is-getting-better-at-understanding-your-code/#how-github-copilot-understands-your-code) | **Details** ([like data from your open files and code you’ve written before and after the cursor](https://github.blog/2023-05-17-how-github-copilot-is-getting-better-at-understanding-your-code/#how-github-copilot-understands-your-code)) **that algorithms send** to a large language model (LLM) as additional information about the code |

## [3 best practices for prompt crafting with GitHub Copilot](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#3-best-practices-for-prompt-crafting-with-github-copilot)

### 1. Set the stage with a high-level goal. 

This is most helpful if you have a blank file or empty codebase. In other words, if GitHub Copilot has zero context of what you want to build or accomplish, setting the stage for the AI pair programmer can be really useful. It helps to prime GitHub Copilot with a big picture description of what you want it to generate—before you jump in with the details.

Think of the process as having a conversation with someone: _How should I break down the problem so we can solve it together? How would I approach pair programming with this person?_

For example, when building a markdown editor in Next.js, we could write a comment like this

```css
/*
Create a basic markdown editor in Next.js with the following features:
- Use react hooks
- Create state for markdown with default text "type markdown here"
- A text area where users can write markdown
- Show a live preview of the markdown text as I type
- Support for basic markdown syntax like headers, bold, italics
- Use React markdown npm package
- The markdown text and resulting HTML should be saved in the component's state and updated in real time
*/
```

This will prompt GitHub Copilot to generate the following code and produce a very simple, unstyled but functional markdown editor in less than 30 seconds. We can use the remaining time to style the component:

![We used this prompt to build a markdown editor in Next.jst using GitHub Copilot: (1) Use react hooks, (2) Create state for markdown with default text 'type markdown here', (3) A text area where users can write markdown, (4) Show a live preview of the markdown text as I type, (5) Support for basic markdown syntax like headers, bold, italics, (6) Use React markdown npm package, (7) The markdown text and resulting HTML should be saved in the component's state and updated in real time.](https://github.blog/wp-content/uploads/2023/06/245577032-c1ee10ef-3c24-458d-94af-8408484a1ee0.gif?resize=1024%2C538)

_Note: this level of detail helps you to create a more desired output, but the results may still be non-deterministic. For example, in the comment, we prompted GitHub Copilot to create default text that says “type markdown here” but instead it generated “markdown preview” as the default words._

### 2. Make your ask simple and specific. Aim to receive a short output from GitHub Copilot. 

Once you communicate your main goal to the AI pair programmer, **articulate the logic and steps it needs to follow for achieving that goal**. GitHub Copilot better understands your goal when you break things down. (Imagine you’re writing a recipe. You’d break the cooking process down into discrete steps–not write a paragraph describing the dish you want to make.)

**Let GitHub Copilot generate the code after each step**, rather than asking it to generate a bunch of code all at once.

Here’s an example of us providing GitHub Copilot with step-by-step instructions for reversing a function:

![We prompted GitHub Copilot to reverse a sentence by writing six prompts one at a time. This allowed GitHub Copilot to generate a suggestion for one prompt before moving onto the text. It also gave us the chance to tweak the suggested code before moving onto the next step. The six prompts we used were: First, let's make the first letter of the sentence lower case if it's not an 'I.' Next, let's split the sentence into an array of words. Then, let's take out the punctuation marks from the sentence. Now, let's remove the punctuation marks from the sentence. Let's reverse the sentence and join it back together. Finally, let's make the first letter of the sentence capital and add the punctuation marks.](https://github.blog/wp-content/uploads/2023/06/245886121-885cc0ed-97c2-4b25-be74-02e7b34711a7.gif)

### 3. Give GitHub Copilot an example or two. 

Learning from examples is not only useful for humans, but also for your AI pair programmer. For instance, we wanted to extract the names from the array of data below and store it in a new array:

```yaml
const data = [
[
{ name: 'John', age: 25 },
{ name: 'Jane', age: 30 }
],
[
{ name: 'Bob', age: 40 }
]
];
```

When we didn’t show GitHub Copilot an example …

```javascript
// Map through an array of arrays of objects to transform data
const data = [
[
{ name: 'John', age: 25 },
{ name: 'Jane', age: 30 }
],
[
{ name: 'Bob', age: 40 }
]
];

const mappedData = data.map(x => [x.name](http://x.name/));

console.log(mappedData);

// Results: [undefined, undefined]
```

It generated an incorrect usage of map:

```javascript
const mappedData = data.map(x => [x.name](http://x.name/));

console.log(mappedData);

// Results: [undefined, undefined]
```

By contrast, when we did provide an example …

```javascript
// Map through an array of arrays of objects
// Example: Extract names from the data array
// Desired outcome: ['John', 'Jane', 'Bob']
const data = [
[{ name: 'John', age: 25 }, { name: 'Jane', age: 30 }],
[{ name: 'Bob', age: 40 }]
];

const mappedData = data.flatMap(sublist => sublist.map(person => person.name));

console.log(mappedData);
```

We received our desired outcome.

```javascript
const mappedData = data.flatMap(sublist => sublist.map(person => person.name));

console.log(mappedData);
// Results: ['John', 'Jane', 'Bob']
```

_Read more about common approaches to AI training, such as [zero-shot, one-shot, and few-shot learning](https://dev.to/github/a-beginners-guide-to-prompt-engineering-with-github-copilot-3ibp/#1424065)._

## [Three additional tips for prompt crafting with GitHub Copilot](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#three-additional-tips-for-prompt-crafting-with-github-copilot)

Here are three additional tips to help guide your conversation with GitHub Copilot.

### [1. Experiment with your prompts.](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#1-experiment-with-your-prompts)

Just how conversation is more of an art than a science, so is prompt crafting. So, if you don’t receive what you want on the first try, recraft your prompt by following the best practices above.

For example, the prompt below is vague. It doesn’t provide any context or boundaries for GitHub Copilot to generate relevant suggestions.

```graphql
# Write some code for grades.py
```

We iterated on the prompt to be more specific, but we still didn’t get the exact result we were looking for. This is a good reminder that adding specificity to your prompt is harder than it sounds. It’s difficult to know, from the start, which details you should include about your goal to generate the most useful suggestions from GitHub Copilot. That’s why we encourage experimentation.

The version of the prompt below is more specific than the one above, but it doesn’t clearly define the input and output requirements.

```graphql
# Implement a function in grades.py to calculate the average grade
```

We experimented with the prompt once more by setting boundaries and outlining what we wanted the function to do. We also rephrased the comment so the function was more clear (giving GitHub Copilot a clear intention to verify against).

This time, we got the results we were looking for.

```graphql
# Implement the function calculate_average_grade in grades.py that takes a list of grades as input and returns the average grade as a floating-point number
```

### [2. Keep a couple of relevant tabs open.](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#2-keep-a-couple-of-relevant-tabs-open)

We don’t have an exact number of tabs that you should keep open to help GitHub Copilot contextualize your code, but from our experience, we’ve found that one or two is helpful.

[GitHub Copilot uses a technique called neighboring tabs](https://github.blog/2023-05-17-how-github-copilot-is-getting-better-at-understanding-your-code/#how-github-copilot-understands-your-code) that allows the AI pair programmer to contextualize your code by processing all of the files open in your IDE instead of just the single file you’re working on. However, it’s not guaranteed that GitHub Copilot will deem all open files as necessary context for your code.

### [3. Use good coding practices.](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#3-use-good-coding-practices)

That includes providing descriptive variable names and functions, and following consistent coding styles and patterns. We’ve found that working with GitHub Copilot encourages us to follow good coding practices we’ve learned throughout our careers.

For example, here we used a descriptive function name and followed the codebase’s patterns of leveraging snake case.

```python
def authenticate_user(username, password):
```

As a result, GitHub Copilot generated a relevant code suggestion:

```python
def authenticate_user(username, password):
# Code for authenticating the user
if is_valid_user(username, password):
generate_session_token(username)
return True
else:
return False
```

Compare this to the example below, where we introduced an inconsistent coding style and poorly named our function.

```python
def rndpwd(l):
```

Instead of suggesting code, GitHub Copilot generated a comment that said, “Code goes here.”

```python
def rndpwd(l):
# Code goes here
```

### [Stay smart](https://github.blog/developer-skills/github/how-to-write-better-prompts-for-github-copilot/?ref_product=copilot&ref_type=engagement&ref_style=text#stay-smart)

The LLMs behind generative AI coding tools are designed to find and extrapolate patterns from their training data, apply those patterns to existing language, and then produce code that follows those patterns. Given the sheer scale of these models, they might generate a code sequence that doesn’t even exist yet. **Just as you would review a colleague’s code, you should always assess, analyze, and validate AI-generated code.**
