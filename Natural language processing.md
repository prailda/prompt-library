---
id: e175b033-c0d0-48cf-b66f-f3ac9349e875
title: Natural language processing (NLP)
tags:
  - nlp
  - natural-language-processing
  - wikipedia
url: https://en.wikipedia.org/wiki/Natural_language_processing
language: en
---
# Natural language processing (NLP)


**Natural language processing** (**NLP**) is the processing of [natural language](https://en.wikipedia.org/wiki/Natural_language "Natural language") information by a [computer](https://en.wikipedia.org/wiki/Computer "Computer"). The study of NLP, a subfield of [computer science](https://en.wikipedia.org/wiki/Computer_science "Computer science"), is generally associated with [artificial intelligence](https://en.wikipedia.org/wiki/Artificial_intelligence "Artificial intelligence"). NLP is related to [information retrieval](https://en.wikipedia.org/wiki/Information_retrieval "Information retrieval"), [knowledge representation](https://en.wikipedia.org/wiki/Knowledge_representation "Knowledge representation"), [computational linguistics](https://en.wikipedia.org/wiki/Computational_linguistics "Computational linguistics"), and more broadly with [linguistics](https://en.wikipedia.org/wiki/Linguistics "Linguistics").[[1]](#cite_note-nlpintro-1)

Major processing tasks in an NLP system include: [speech recognition](https://en.wikipedia.org/wiki/Speech_recognition "Speech recognition"), [text classification](https://en.wikipedia.org/wiki/Text_classification "Text classification"), [natural language understanding](https://en.wikipedia.org/wiki/Natural-language_understanding "Natural-language understanding"), and [natural language generation](https://en.wikipedia.org/wiki/Natural_language_generation "Natural language generation").

Natural language processing has its roots in the 1950s.[[2]](#cite_note-2) Already in 1950, [Alan Turing](https://en.wikipedia.org/wiki/Alan_Turing "Alan Turing") published an article titled "[Computing Machinery and Intelligence](https://en.wikipedia.org/wiki/Computing_Machinery_and_Intelligence "Computing Machinery and Intelligence")" which proposed what is now called the [Turing test](https://en.wikipedia.org/wiki/Turing_test "Turing test") as a criterion of intelligence, though at the time that was not articulated as a problem separate from artificial intelligence. The proposed test includes a task that involves the automated interpretation and generation of natural language.

### Symbolic NLP (1950s – early 1990s)

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=2 "Edit section: Symbolic NLP (1950s – early 1990s)")]

[![](https://upload.wikimedia.org/wikipedia/commons/thumb/0/09/AST_Document.svg/250px-AST_Document.svg.png)](https://en.wikipedia.org/wiki/File:AST_Document.svg)

A document parsed into an abstract syntax tree

The premise of symbolic NLP is well-summarized by [John Searle](https://en.wikipedia.org/wiki/John_Searle "John Searle")'s [Chinese room](https://en.wikipedia.org/wiki/Chinese_room "Chinese room") experiment: Given a collection of rules (e.g., a Chinese phrasebook, with questions and matching answers), the computer emulates natural language understanding (or other NLP tasks) by applying those rules to the data it confronts.

- **1950s**: The [Georgetown experiment](https://en.wikipedia.org/wiki/Georgetown-IBM_experiment "Georgetown-IBM experiment") in 1954 involved fully [automatic translation](https://en.wikipedia.org/wiki/Automatic_translation "Automatic translation") of more than sixty Russian sentences into English. The authors claimed that within three or five years, machine translation would be a solved problem.[[3]](#cite_note-3) However, real progress was much slower, and after the [ALPAC report](https://en.wikipedia.org/wiki/ALPAC "ALPAC") in 1966, which found that ten years of research had failed to fulfill the expectations, funding for machine translation was dramatically reduced. Little further research in machine translation was conducted in America (though some research continued elsewhere, such as Japan and Europe[[4]](#cite_note-4)) until the late 1980s when the first [statistical machine translation](https://en.wikipedia.org/wiki/Statistical_machine_translation "Statistical machine translation") systems were developed.
- **1960s**: Some notably successful natural language processing systems developed in the 1960s were [SHRDLU](https://en.wikipedia.org/wiki/SHRDLU "SHRDLU"), a natural language system working in restricted "[blocks worlds](https://en.wikipedia.org/wiki/Blocks_world "Blocks world")" with restricted vocabularies, and [ELIZA](https://en.wikipedia.org/wiki/ELIZA "ELIZA"), a simulation of a [Rogerian psychotherapist](https://en.wikipedia.org/wiki/Rogerian_psychotherapy "Rogerian psychotherapy"), written by [Joseph Weizenbaum](https://en.wikipedia.org/wiki/Joseph_Weizenbaum "Joseph Weizenbaum") between 1964 and 1966. Using almost no information about human thought or emotion, ELIZA sometimes provided a startlingly human-like interaction. When the "patient" exceeded the very small knowledge base, ELIZA might provide a generic response, for example, responding to "My head hurts" with "Why do you say your head hurts?". Ross Quillian's successful work on natural language was demonstrated with a vocabulary of only _twenty_ words, because that was all that would fit in a computer memory at the time.[[5]](#cite_note-5)

- **1970s**: During the 1970s, many programmers began to write "conceptual [ontologies](https://en.wikipedia.org/wiki/Ontology_\(information_science\) "Ontology (information science)")", which structured real-world information into computer-understandable data. Examples are MARGIE (Schank, 1975), SAM (Cullingford, 1978), PAM (Wilensky, 1978), TaleSpin (Meehan, 1976), QUALM (Lehnert, 1977), Politics (Carbonell, 1979), and Plot Units (Lehnert 1981). During this time, the first [chatterbots](https://en.wikipedia.org/wiki/Chatterbots "Chatterbots") were written (e.g., [PARRY](https://en.wikipedia.org/wiki/PARRY "PARRY")).
- **1980s**: The 1980s and early 1990s mark the heyday of symbolic methods in NLP. Focus areas of the time included research on rule-based parsing (e.g., the development of [HPSG](https://en.wikipedia.org/wiki/Head-driven_phrase_structure_grammar "Head-driven phrase structure grammar") as a computational operationalization of [generative grammar](https://en.wikipedia.org/wiki/Generative_grammar "Generative grammar")), morphology (e.g., two-level morphology[[6]](#cite_note-6)), semantics (e.g., [Lesk algorithm](https://en.wikipedia.org/wiki/Lesk_algorithm "Lesk algorithm")), reference (e.g., within Centering Theory[[7]](#cite_note-7)) and other areas of natural language understanding (e.g., in the [Rhetorical Structure Theory](https://en.wikipedia.org/wiki/Rhetorical_structure_theory "Rhetorical structure theory")). Other lines of research were continued, e.g., the development of chatterbots with [Racter](https://en.wikipedia.org/wiki/Racter "Racter") and [Jabberwacky](https://en.wikipedia.org/wiki/Jabberwacky "Jabberwacky"). An important development (that eventually led to the statistical turn in the 1990s) was the rising importance of quantitative evaluation in this period.[[8]](#cite_note-8)

### Statistical NLP (1990s–present)

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=3 "Edit section: Statistical NLP (1990s–present)")]

Up until the 1980s, most natural language processing systems were based on complex sets of hand-written rules. Starting in the late 1980s, however, there was a revolution in natural language processing with the introduction of [machine learning](https://en.wikipedia.org/wiki/Machine_learning "Machine learning") algorithms for language processing. This was due to both the steady increase in computational power (see [Moore's law](https://en.wikipedia.org/wiki/Moore%27s_law "Moore's law")) and the gradual lessening of the dominance of [Chomskyan](https://en.wikipedia.org/wiki/Noam_Chomsky "Noam Chomsky") theories of linguistics (e.g. [transformational grammar](https://en.wikipedia.org/wiki/Transformational_grammar "Transformational grammar")), whose theoretical underpinnings discouraged the sort of [corpus linguistics](https://en.wikipedia.org/wiki/Corpus_linguistics "Corpus linguistics") that underlies the machine-learning approach to language processing.[[9]](#cite_note-9)

- **1990s**: Many of the notable early successes in statistical methods in NLP occurred in the field of [machine translation](https://en.wikipedia.org/wiki/Machine_translation "Machine translation"), due especially to work at IBM Research, such as [IBM alignment models](https://en.wikipedia.org/wiki/IBM_alignment_models "IBM alignment models"). These systems were able to take advantage of existing multilingual [textual corpora](https://en.wikipedia.org/wiki/Text_corpus "Text corpus") that had been produced by the [Parliament of Canada](https://en.wikipedia.org/wiki/Parliament_of_Canada "Parliament of Canada") and the [European Union](https://en.wikipedia.org/wiki/European_Union "European Union") as a result of laws calling for the translation of all governmental proceedings into all official languages of the corresponding systems of government. However, most other systems depended on corpora specifically developed for the tasks implemented by these systems, which was (and often continues to be) a major limitation in the success of these systems. As a result, a great deal of research has gone into methods of more effectively learning from limited amounts of data.
- **2000s**: With the growth of the web, increasing amounts of raw (unannotated) language data have become available since the mid-1990s. Research has thus increasingly focused on [unsupervised](https://en.wikipedia.org/wiki/Unsupervised_learning "Unsupervised learning") and [semi-supervised learning](https://en.wikipedia.org/wiki/Semi-supervised_learning "Semi-supervised learning") algorithms. Such algorithms can learn from data that has not been hand-annotated with the desired answers or using a combination of annotated and non-annotated data. Generally, this task is much more difficult than [supervised learning](https://en.wikipedia.org/wiki/Supervised_learning "Supervised learning"), and typically produces less accurate results for a given amount of input data. However, there is an enormous amount of non-annotated data available (including, among other things, the entire content of the [World Wide Web](https://en.wikipedia.org/wiki/World_Wide_Web "World Wide Web")), which can often make up for the worse efficiency if the algorithm used has a low enough [time complexity](https://en.wikipedia.org/wiki/Time_complexity "Time complexity") to be practical.
- **2003:** [word n-gram model](https://en.wikipedia.org/wiki/Word_n-gram_language_model "Word n-gram language model"), at the time the best statistical algorithm, is outperformed by a [multi-layer perceptron](https://en.wikipedia.org/wiki/Multi-layer_perceptron "Multi-layer perceptron") (with a single hidden layer and context length of several words, trained on up to 14 million words, by [Bengio](https://en.wikipedia.org/wiki/Yoshua_Bengio "Yoshua Bengio") et al.)[[10]](#cite_note-10)
- **2010:** [Tomáš Mikolov](https://en.wikipedia.org/wiki/Tom%C3%A1%C5%A1_Mikolov "Tomáš Mikolov") (then a PhD student at [Brno University of Technology](https://en.wikipedia.org/wiki/Brno_University_of_Technology "Brno University of Technology")) with co-authors applied a simple [recurrent neural network](https://en.wikipedia.org/wiki/Recurrent_neural_network "Recurrent neural network") with a single hidden layer to language modelling,[[11]](#cite_note-11) and in the following years he went on to develop [Word2vec](https://en.wikipedia.org/wiki/Word2vec "Word2vec"). In the 2010s, [representation learning](https://en.wikipedia.org/wiki/Representation_learning "Representation learning") and [deep neural network](https://en.wikipedia.org/wiki/Deep_learning "Deep learning")-style (featuring many hidden layers) machine learning methods became widespread in natural language processing. That popularity was due partly to a flurry of results showing that such techniques[[12]](#cite_note-goldberg:nnlp17-12)[[13]](#cite_note-goodfellow:book16-13) can achieve state-of-the-art results in many natural language tasks, e.g., in [language modeling](https://en.wikipedia.org/wiki/Language_modeling "Language modeling")[[14]](#cite_note-jozefowicz:lm16-14) and parsing.[[15]](#cite_note-choe:emnlp16-15)[[16]](#cite_note-vinyals:nips15-16) This is increasingly important [in medicine and healthcare](https://en.wikipedia.org/wiki/Artificial_intelligence_in_healthcare "Artificial intelligence in healthcare"), where NLP helps analyze notes and text in [electronic health records](https://en.wikipedia.org/wiki/Electronic_health_record "Electronic health record") that would otherwise be inaccessible for study when seeking to improve care[[17]](#cite_note-17) or protect patient privacy.[[18]](#cite_note-18)

## Approaches: Symbolic, statistical, neural networks

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=4 "Edit section: Approaches: Symbolic, statistical, neural networks")]

Symbolic approach, i.e., the hand-coding of a set of rules for manipulating symbols, coupled with a dictionary lookup, was historically the first approach used both by AI in general and by NLP in particular:[[19]](#cite_note-winograd:shrdlu71-19)[[20]](#cite_note-schank77-20) such as by writing grammars or devising heuristic rules for [stemming](https://en.wikipedia.org/wiki/Stemming "Stemming").

[Machine learning](https://en.wikipedia.org/wiki/Machine_learning "Machine learning") approaches, which include both statistical and neural networks, on the other hand, have many advantages over the symbolic approach:

- both statistical and neural networks methods can focus more on the most common cases extracted from a corpus of texts, whereas the rule-based approach needs to provide rules for both rare cases and common ones equally.

- [language models](https://en.wikipedia.org/wiki/Language_model "Language model"), produced by either statistical or neural networks methods, are more robust to both unfamiliar (e.g. containing words or structures that have not been seen before) and erroneous input (e.g. with misspelled words or words accidentally omitted) in comparison to the rule-based systems, which are also more costly to produce.

- the larger such a (probabilistic) language model is, the more accurate it becomes, in contrast to rule-based systems that can gain accuracy only by increasing the amount and complexity of the rules leading to [intractability](https://en.wikipedia.org/wiki/Intractable_problem "Intractable problem") problems.

Rule-based systems are commonly used:

- when the amount of training data is insufficient to successfully apply machine learning methods, e.g., for the machine translation of low-resource languages such as provided by the [Apertium](https://en.wikipedia.org/wiki/Apertium "Apertium") system,
- for preprocessing in NLP pipelines, e.g., [tokenization](https://en.wikipedia.org/wiki/Tokenization_\(lexical_analysis\) "Tokenization (lexical analysis)"), or
- for postprocessing and transforming the output of NLP pipelines, e.g., for [knowledge extraction](https://en.wikipedia.org/wiki/Knowledge_extraction "Knowledge extraction") from syntactic parses.

### Statistical approach

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=5 "Edit section: Statistical approach")]

In the late 1980s and mid-1990s, the statistical approach ended a period of [AI winter](https://en.wikipedia.org/wiki/AI_winter "AI winter"), which was caused by the inefficiencies of the rule-based approaches.[[21]](#cite_note-johnson:eacl:ilcl09-21)[[22]](#cite_note-resnik:langlog11-22)

The earliest [decision trees](https://en.wikipedia.org/wiki/Decision_tree "Decision tree"), producing systems of hard [if–then rules](https://en.wikipedia.org/wiki/Conditional_\(computer_programming\)#If%E2%80%93then\(%E2%80%93else\) "Conditional (computer programming)"), were still very similar to the old rule-based approaches. Only the introduction of hidden [Markov models](https://en.wikipedia.org/wiki/Markov_model "Markov model"), applied to part-of-speech tagging, announced the end of the old rule-based approach.

A major drawback of statistical methods is that they require elaborate [feature engineering](https://en.wikipedia.org/wiki/Feature_engineering "Feature engineering"). Since 2015,[[23]](#cite_note-23) the statistical approach has been replaced by the [neural networks](https://en.wikipedia.org/wiki/Artificial_neural_network "Artificial neural network") approach, using [semantic networks](https://en.wikipedia.org/wiki/Semantic_networks "Semantic networks")[[24]](#cite_note-24) and [word embeddings](https://en.wikipedia.org/wiki/Word_embedding "Word embedding") to capture semantic properties of words.

Intermediate tasks (e.g., part-of-speech tagging and dependency parsing) are not needed anymore.

[Neural machine translation](https://en.wikipedia.org/wiki/Neural_machine_translation "Neural machine translation"), based on then-newly invented [sequence-to-sequence](https://en.wikipedia.org/wiki/Seq2seq "Seq2seq") transformations, made obsolete the intermediate steps, such as word alignment, previously necessary for [statistical machine translation](https://en.wikipedia.org/wiki/Statistical_machine_translation "Statistical machine translation").

The following is a list of some of the most commonly researched tasks in natural language processing. Some of these tasks have direct real-world applications, while others more commonly serve as subtasks that are used to aid in solving larger tasks.

Though natural language processing tasks are closely intertwined, they can be subdivided into categories for convenience. A coarse division is given below.

### Text and speech processing

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=8 "Edit section: Text and speech processing")]

[![](https://upload.wikimedia.org/wikipedia/commons/thumb/8/8c/Hebrew_stop_words_word_cloud.png/250px-Hebrew_stop_words_word_cloud.png)](https://en.wikipedia.org/wiki/File:Hebrew_stop_words_word_cloud.png)

Word cloud of stop words in Hebrew

[Optical character recognition](https://en.wikipedia.org/wiki/Optical_character_recognition "Optical character recognition") (OCR)

Given an image representing printed text, determine the corresponding text.

[Speech recognition](https://en.wikipedia.org/wiki/Speech_recognition "Speech recognition")

Given a sound clip of a person or people speaking, determine the textual representation of the speech. This is the opposite of [text to speech](https://en.wikipedia.org/wiki/Text_to_speech "Text to speech") and is one of the extremely difficult problems colloquially termed "[AI-complete](https://en.wikipedia.org/wiki/AI-complete "AI-complete")" (see above). In [natural speech](https://en.wikipedia.org/wiki/Natural_speech "Natural speech") there are hardly any pauses between successive words, and thus [speech segmentation](https://en.wikipedia.org/wiki/Speech_segmentation "Speech segmentation") is a necessary subtask of speech recognition (see below). In most spoken languages, the sounds representing successive letters blend into each other in a process termed [coarticulation](https://en.wikipedia.org/wiki/Coarticulation "Coarticulation"), so the conversion of the [analog signal](https://en.wikipedia.org/wiki/Analog_signal "Analog signal") to discrete characters can be a very difficult process. Also, given that words in the same language are spoken by people with different accents, the speech recognition software must be able to recognize the wide variety of input as being identical to each other in terms of its textual equivalent.

[Speech segmentation](https://en.wikipedia.org/wiki/Speech_segmentation "Speech segmentation")

Given a sound clip of a person or people speaking, separate it into words. A subtask of [speech recognition](https://en.wikipedia.org/wiki/Speech_recognition "Speech recognition") and typically grouped with it.

[Text-to-speech](https://en.wikipedia.org/wiki/Text-to-speech "Text-to-speech")

Given a text, transform those units and produce a spoken representation. Text-to-speech can be used to aid the visually impaired.[[25]](#cite_note-25)

[Word segmentation](https://en.wikipedia.org/wiki/Word_segmentation "Word segmentation") ([Tokenization](https://en.wikipedia.org/wiki/Tokenization_\(lexical_analysis\) "Tokenization (lexical analysis)"))

Tokenization is a process used in text analysis that divides text into individual words or word fragments. This technique results in two key components: a word index and tokenized text. The word index is a list that maps unique words to specific numerical identifiers, and the tokenized text replaces each word with its corresponding numerical token. These numerical tokens are then used in various deep learning methods.[[26]](#cite_note-:0-26)

For a language like [English](https://en.wikipedia.org/wiki/English_language "English language"), this is fairly trivial, since words are usually separated by spaces. However, some written languages like [Chinese](https://en.wikipedia.org/wiki/Chinese_language "Chinese language"), [Japanese](https://en.wikipedia.org/wiki/Japanese_language "Japanese language") and [Thai](https://en.wikipedia.org/wiki/Thai_language "Thai language") do not mark word boundaries in such a fashion, and in those languages text segmentation is a significant task requiring knowledge of the [vocabulary](https://en.wikipedia.org/wiki/Vocabulary "Vocabulary") and [morphology](https://en.wikipedia.org/wiki/Morphology_\(linguistics\) "Morphology (linguistics)") of words in the language. Sometimes this process is also used in cases like [bag of words](https://en.wikipedia.org/wiki/Bag_of_words "Bag of words") (BOW) creation in data mining.[_[citation needed](https://en.wikipedia.org/wiki/Wikipedia:Citation_needed "Wikipedia:Citation needed")_]

### Morphological analysis

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=9 "Edit section: Morphological analysis")]

[![](https://upload.wikimedia.org/wikipedia/commons/thumb/2/27/Eustagger_euskal_lematizatzailearen_adibidea.png/250px-Eustagger_euskal_lematizatzailearen_adibidea.png)](https://en.wikipedia.org/wiki/File:Eustagger_euskal_lematizatzailearen_adibidea.png)

Lemmatization of Basque words

[Lemmatization](https://en.wikipedia.org/wiki/Lemmatisation "Lemmatisation")

The task of removing inflectional endings only and to return the base dictionary form of a word which is also known as a lemma. Lemmatization is another technique for reducing words to their normalized form. But in this case, the transformation actually uses a dictionary to map words to their actual form.[[27]](#cite_note-27)

[Morphological segmentation](https://en.wikipedia.org/wiki/Morphology_\(linguistics\) "Morphology (linguistics)")

Separate words into individual [morphemes](https://en.wikipedia.org/wiki/Morpheme "Morpheme") and identify the class of the morphemes. The difficulty of this task depends greatly on the complexity of the [morphology](https://en.wikipedia.org/wiki/Morphology_\(linguistics\) "Morphology (linguistics)") (_i.e._, the structure of words) of the language being considered. [English](https://en.wikipedia.org/wiki/English_language "English language") has fairly simple morphology, especially [inflectional morphology](https://en.wikipedia.org/wiki/Inflectional_morphology "Inflectional morphology"), and thus it is often possible to ignore this task entirely and simply model all possible forms of a word (e.g., "open, opens, opened, opening") as separate words. In languages such as [Turkish](https://en.wikipedia.org/wiki/Turkish_language "Turkish language") or [Meitei](https://en.wikipedia.org/wiki/Meitei_language "Meitei language"), a highly [agglutinated](https://en.wikipedia.org/wiki/Agglutination "Agglutination") Indian language, however, such an approach is not possible, as each dictionary entry has thousands of possible word forms.[[28]](#cite_note-28)

[Part-of-speech tagging](https://en.wikipedia.org/wiki/Part-of-speech_tagging "Part-of-speech tagging")

Given a sentence, determine the [part of speech](https://en.wikipedia.org/wiki/Part_of_speech "Part of speech") (POS) for each word. Many words, especially common ones, can serve as multiple parts of speech. For example, "book" can be a [noun](https://en.wikipedia.org/wiki/Noun "Noun") ("the book on the table") or [verb](https://en.wikipedia.org/wiki/Verb "Verb") ("to book a flight"); "set" can be a noun, verb or [adjective](https://en.wikipedia.org/wiki/Adjective "Adjective"); and "out" can be any of at least five different parts of speech.

[Stemming](https://en.wikipedia.org/wiki/Stemming "Stemming")

The process of reducing inflected (or sometimes derived) words to a base form (e.g., "close" will be the root for "closed", "closing", "close", "closer" etc.). Stemming yields similar results as lemmatization, but does so on grounds of rules, not a dictionary.

[Grammar induction](https://en.wikipedia.org/wiki/Grammar_induction "Grammar induction")[[29]](#cite_note-29)

Generate a [formal grammar](https://en.wikipedia.org/wiki/Formal_grammar "Formal grammar") that describes a language's syntax.

[Sentence breaking](https://en.wikipedia.org/wiki/Sentence_breaking "Sentence breaking") (also known as "[sentence boundary disambiguation](https://en.wikipedia.org/wiki/Sentence_boundary_disambiguation "Sentence boundary disambiguation")")

Given a chunk of text, find the sentence boundaries. Sentence boundaries are often marked by [periods](https://en.wikipedia.org/wiki/Full_stop "Full stop") or other [punctuation marks](https://en.wikipedia.org/wiki/Punctuation_mark "Punctuation mark"), but these same characters can serve other purposes (e.g., marking [abbreviations](https://en.wikipedia.org/wiki/Abbreviation "Abbreviation")).

[Parsing](https://en.wikipedia.org/wiki/Parsing "Parsing")

Determine the [parse tree](https://en.wikipedia.org/wiki/Parse_tree "Parse tree") (grammatical analysis) of a given sentence. The [grammar](https://en.wikipedia.org/wiki/Grammar "Grammar") for [natural languages](https://en.wikipedia.org/wiki/Natural_language "Natural language") is [ambiguous](https://en.wikipedia.org/wiki/Ambiguous "Ambiguous") and typical sentences have multiple possible analyses: perhaps surprisingly, for a typical sentence there may be thousands of potential parses (most of which will seem completely nonsensical to a human). There are two primary types of parsing: _dependency parsing_ and _constituency parsing_. Dependency parsing focuses on the relationships between words in a sentence (marking things like primary objects and predicates), whereas constituency parsing focuses on building out the parse tree using a [probabilistic context-free grammar](https://en.wikipedia.org/wiki/Probabilistic_context-free_grammar "Probabilistic context-free grammar") (PCFG) (see also _[stochastic grammar](https://en.wikipedia.org/wiki/Stochastic_grammar "Stochastic grammar")_).

### Lexical semantics (of individual words in context)

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=11 "Edit section: Lexical semantics (of individual words in context)")]

[![](https://upload.wikimedia.org/wikipedia/commons/thumb/3/34/Entity_Linking_-_Example_of_pipeline.png/250px-Entity_Linking_-_Example_of_pipeline.png)](https://en.wikipedia.org/wiki/File:Entity_Linking_-_Example_of_pipeline.png)

An entity linking pipeline

[Lexical semantics](https://en.wikipedia.org/wiki/Lexical_semantics "Lexical semantics")

What is the computational meaning of individual words in context?

[Distributional semantics](https://en.wikipedia.org/wiki/Distributional_semantics "Distributional semantics")

How can we learn semantic representations from data?

[Named entity recognition](https://en.wikipedia.org/wiki/Named_entity_recognition "Named entity recognition") (NER)

Given a stream of text, determine which items in the text map to proper names, such as people or places, and what the type of each such name is (e.g. person, location, organization). Although [capitalization](https://en.wikipedia.org/wiki/Capitalization "Capitalization") can aid in recognizing named entities in languages such as English, this information cannot aid in determining the type of [named entity](https://en.wikipedia.org/wiki/Named_entity "Named entity"), and in any case, is often inaccurate or insufficient. For example, the first letter of a sentence is also capitalized, and named entities often span several words, only some of which are capitalized. Furthermore, many other languages in non-Western scripts (e.g. [Chinese](https://en.wikipedia.org/wiki/Chinese_language "Chinese language") or [Arabic](https://en.wikipedia.org/wiki/Arabic_language "Arabic language")) do not have any capitalization at all, and even languages with capitalization may not consistently use it to distinguish names. For example, [German](https://en.wikipedia.org/wiki/German_language "German language") capitalizes all [nouns](https://en.wikipedia.org/wiki/Noun "Noun"), regardless of whether they are names, and [French](https://en.wikipedia.org/wiki/French_language "French language") and [Spanish](https://en.wikipedia.org/wiki/Spanish_language "Spanish language") do not capitalize names that serve as [adjectives](https://en.wikipedia.org/wiki/Adjective "Adjective"). Another name for this task is token classification.[[30]](#cite_note-30)

[Sentiment analysis](https://en.wikipedia.org/wiki/Sentiment_analysis "Sentiment analysis") (see also [Multimodal sentiment analysis](https://en.wikipedia.org/wiki/Multimodal_sentiment_analysis "Multimodal sentiment analysis"))

Sentiment analysis is a computational method used to identify and classify the emotional intent behind text. This technique involves analyzing text to determine whether the expressed sentiment is positive, negative, or neutral. Models for sentiment classification typically utilize inputs such as [word n-grams](https://en.wikipedia.org/wiki/Word_n-gram_language_model "Word n-gram language model"), [Term Frequency-Inverse Document Frequency](https://en.wikipedia.org/wiki/Term_frequency-inverse_document_frequency "Term frequency-inverse document frequency") (TF-IDF) features, hand-generated features, or employ [deep learning](https://en.wikipedia.org/wiki/Deep_learning "Deep learning") models designed to recognize both long-term and short-term dependencies in text sequences. The applications of sentiment analysis are diverse, extending to tasks such as categorizing customer reviews on various online platforms.[[26]](#cite_note-:0-26)

[Terminology extraction](https://en.wikipedia.org/wiki/Terminology_extraction "Terminology extraction")

The goal of terminology extraction is to automatically extract relevant terms from a given corpus.

[Word-sense disambiguation](https://en.wikipedia.org/wiki/Word-sense_disambiguation "Word-sense disambiguation") (WSD)

Many words have more than one [meaning](https://en.wikipedia.org/wiki/Meaning_\(linguistics\) "Meaning (linguistics)"); we have to select the meaning which makes the most sense in context. For this problem, we are typically given a list of words and associated word senses, e.g. from a dictionary or an online resource such as [WordNet](https://en.wikipedia.org/wiki/WordNet "WordNet").

[Entity linking](https://en.wikipedia.org/wiki/Entity_linking "Entity linking")

Many words—typically proper names—refer to [named entities](https://en.wikipedia.org/wiki/Named_entity "Named entity"); here we have to select the entity (a famous individual, a location, a company, etc.) which is referred to in context.

### Relational semantics (semantics of individual sentences)

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=12 "Edit section: Relational semantics (semantics of individual sentences)")]

[Relationship extraction](https://en.wikipedia.org/wiki/Relationship_extraction "Relationship extraction")

Given a chunk of text, identify the relationships among named entities (e.g. who is married to whom).

[Semantic parsing](https://en.wikipedia.org/wiki/Semantic_parsing "Semantic parsing")

Given a piece of text (typically a sentence), produce a formal representation of its semantics, either as a graph (e.g., in [AMR parsing](https://en.wikipedia.org/wiki/Abstract_Meaning_Representation "Abstract Meaning Representation")) or in accordance with a logical formalism (e.g., in [DRT parsing](https://en.wikipedia.org/wiki/Discourse_representation_theory "Discourse representation theory")). This challenge typically includes aspects of several more elementary NLP tasks from semantics (e.g., semantic role labelling, word-sense disambiguation) and can be extended to include full-fledged discourse analysis (e.g., discourse analysis, coreference; see [Natural language understanding](#Natural_language_understanding) below).

[Semantic role labelling](https://en.wikipedia.org/wiki/Semantic_role_labeling "Semantic role labeling") (see also implicit semantic role labelling below)

Given a single sentence, identify and disambiguate semantic predicates (e.g., verbal [frames](https://en.wikipedia.org/wiki/Frame_semantics_\(linguistics\) "Frame semantics (linguistics)")), then identify and classify the frame elements ([semantic roles](https://en.wikipedia.org/wiki/Semantic_roles "Semantic roles")).

### Discourse (semantics beyond individual sentences)

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=13 "Edit section: Discourse (semantics beyond individual sentences)")]

[Coreference resolution](https://en.wikipedia.org/wiki/Coreference "Coreference")

Given a sentence or larger chunk of text, determine which words ("mentions") refer to the same objects ("entities"). [Anaphora resolution](https://en.wikipedia.org/wiki/Anaphora_resolution "Anaphora resolution") is a specific example of this task, and is specifically concerned with matching up [pronouns](https://en.wikipedia.org/wiki/Pronoun "Pronoun") with the nouns or names to which they refer. The more general task of coreference resolution also includes identifying so-called "bridging relationships" involving [referring expressions](https://en.wikipedia.org/wiki/Referring_expression "Referring expression"). For example, in a sentence such as "He entered John's house through the front door", "the front door" is a referring expression and the bridging relationship to be identified is the fact that the door being referred to is the front door of John's house (rather than of some other structure that might also be referred to).

[Discourse analysis](https://en.wikipedia.org/wiki/Discourse_analysis "Discourse analysis")

This rubric includes several related tasks. One task is discourse parsing, i.e., identifying the [discourse](https://en.wikipedia.org/wiki/Discourse "Discourse") structure of a connected text, i.e. the nature of the discourse relationships between sentences (e.g. elaboration, explanation, contrast). Another possible task is recognizing and classifying the [speech acts](https://en.wikipedia.org/wiki/Speech_act "Speech act") in a chunk of text (e.g. yes–no question, content question, statement, assertion, etc.).

Implicit semantic role labelling

Given a single sentence, identify and disambiguate semantic predicates (e.g., verbal [frames](https://en.wikipedia.org/wiki/Frame_semantics_\(linguistics\) "Frame semantics (linguistics)")) and their explicit semantic roles in the current sentence (see [Semantic role labelling](#Semantic_role_labelling) above). Then, identify semantic roles that are not explicitly realized in the current sentence, classify them into arguments that are explicitly realized elsewhere in the text and those that are not specified, and resolve the former against the local text. A closely related task is zero anaphora resolution, i.e., the extension of coreference resolution to [pro-drop languages](https://en.wikipedia.org/wiki/Pro-drop_language "Pro-drop language").

[Recognizing textual entailment](https://en.wikipedia.org/wiki/Textual_entailment "Textual entailment")

Given two text fragments, determine if one being true entails the other, entails the other's negation, or allows the other to be either true or false.[[31]](#cite_note-rte:11-31)

[Topic segmentation](https://en.wikipedia.org/wiki/Topic_segmentation "Topic segmentation") and recognition

Given a chunk of text, separate it into segments each of which is devoted to a topic, and identify the topic of the segment.

[Argument mining](https://en.wikipedia.org/wiki/Argument_mining "Argument mining")

The goal of argument mining is the automatic extraction and identification of argumentative structures from [natural language](https://en.wikipedia.org/wiki/Natural_language "Natural language") text with the aid of computer programs.[[32]](#cite_note-32) Such argumentative structures include the premise, conclusions, the [argument scheme](https://en.wikipedia.org/wiki/Argument_scheme "Argument scheme") and the relationship between the main and subsidiary argument, or the main and counter-argument within discourse.[[33]](#cite_note-33)[[34]](#cite_note-34)

### Higher-level NLP applications

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=14 "Edit section: Higher-level NLP applications")]

[![](https://upload.wikimedia.org/wikipedia/commons/thumb/4/43/Use_of_Firefox%27_%27Translate_selection_to_English%27_machine_translation_on_Commons_talk_meta_pages_02.png/250px-Use_of_Firefox%27_%27Translate_selection_to_English%27_machine_translation_on_Commons_talk_meta_pages_02.png)](https://en.wikipedia.org/wiki/File:Use_of_Firefox%27_%27Translate_selection_to_English%27_machine_translation_on_Commons_talk_meta_pages_02.png)

Machine translation in Firefox

[Automatic summarization](https://en.wikipedia.org/wiki/Automatic_summarization "Automatic summarization") (text summarization)

Produce a readable summary of a chunk of text. Often used to provide summaries of the text of a known type, such as research papers, articles in the financial section of a newspaper.

Grammatical error correction

Grammatical error detection and correction involves a great band-width of problems on all levels of linguistic analysis (phonology/orthography, morphology, syntax, semantics, pragmatics). Grammatical error correction is impactful since it affects hundreds of millions of people that use or acquire English as a second language. It has thus been subject to a number of shared tasks since 2011.[[35]](#cite_note-35)[[36]](#cite_note-36)[[37]](#cite_note-37) As far as orthography, morphology, syntax and certain aspects of semantics are concerned, and due to the development of powerful neural language models such as [GPT-2](https://en.wikipedia.org/wiki/GPT-2 "GPT-2"), this can now (2019) be considered a largely solved problem and is being marketed in various commercial applications.

[Logic translation](https://en.wikipedia.org/wiki/Logic_translation "Logic translation")

Translate a text from a natural language into formal logic.

[Machine translation](https://en.wikipedia.org/wiki/Machine_translation "Machine translation") (MT)

Automatically translate text from one human language to another. This is one of the most difficult problems, and is a member of a class of problems colloquially termed "[AI-complete](https://en.wikipedia.org/wiki/AI-complete "AI-complete")", i.e. requiring all of the different types of knowledge that humans possess (grammar, semantics, facts about the real world, etc.) to solve properly.

[Natural language understanding](https://en.wikipedia.org/wiki/Natural_language_understanding "Natural language understanding") (NLU)

Convert chunks of text into more formal representations such as [first-order logic](https://en.wikipedia.org/wiki/First-order_logic "First-order logic") structures that are easier for [computer](https://en.wikipedia.org/wiki/Computer "Computer") programs to manipulate. Natural language understanding involves the identification of the intended semantic from the multiple possible semantics which can be derived from a natural language expression which usually takes the form of organized notations of natural language concepts. Introduction and creation of language metamodel and ontology are efficient however empirical solutions. An explicit formalization of natural language semantics without confusions with implicit assumptions such as [closed-world assumption](https://en.wikipedia.org/wiki/Closed-world_assumption "Closed-world assumption") (CWA) vs. [open-world assumption](https://en.wikipedia.org/wiki/Open-world_assumption "Open-world assumption"), or subjective Yes/No vs. objective True/False is expected for the construction of a basis of semantics formalization.[[38]](#cite_note-38)

[Natural language generation](https://en.wikipedia.org/wiki/Natural_language_generation "Natural language generation") (NLG):

Convert information from computer databases or semantic intents into readable human language.

Book generation

Not an NLP task proper but an extension of natural language generation and other NLP tasks is the creation of full-fledged books. The first machine-generated book was created by a rule-based system in 1984 (Racter, _The policeman's beard is half-constructed_).[[39]](#cite_note-39) The first published work by a neural network was published in 2018, _[1 the Road](https://en.wikipedia.org/wiki/1_the_Road "1 the Road")_, marketed as a novel, contains sixty million words. Both these systems are basically elaborate but non-sensical (semantics-free) [language models](https://en.wikipedia.org/wiki/Language_model "Language model"). The first machine-generated science book was published in 2019 (Beta Writer, _Lithium-Ion Batteries_, Springer, Cham).[[40]](#cite_note-40) Unlike _Racter_ and _1 the Road_, this is grounded on factual knowledge and based on text summarization.

[Document AI](https://en.wikipedia.org/wiki/Document_AI "Document AI")

A Document AI platform sits on top of the NLP technology enabling users with no prior experience of artificial intelligence, machine learning or NLP to quickly train a computer to extract the specific data they need from different document types. NLP-powered Document AI enables non-technical teams to quickly access information hidden in documents, for example, lawyers, business analysts and accountants.[[41]](#cite_note-41)

[Dialogue management](https://en.wikipedia.org/wiki/Dialogue_system "Dialogue system")

Computer systems intended to converse with a human.

[Question answering](https://en.wikipedia.org/wiki/Question_answering "Question answering")

Given a human-language question, determine its answer. Typical questions have a specific right answer (such as "What is the capital of Canada?"), but sometimes open-ended questions are also considered (such as "What is the meaning of life?").

[Text-to-image generation](https://en.wikipedia.org/wiki/Text-to-image_generation "Text-to-image generation")

Given a description of an image, generate an image that matches the description.[[42]](#cite_note-42)

Text-to-scene generation

Given a description of a scene, generate a [3D model](https://en.wikipedia.org/wiki/3D_model "3D model") of the scene.[[43]](#cite_note-43)[[44]](#cite_note-44)

[Text-to-video](https://en.wikipedia.org/wiki/Text-to-video_model "Text-to-video model")

Given a description of a video, generate a video that matches the description.[[45]](#cite_note-45)[[46]](#cite_note-46)

## General tendencies and (possible) future directions

[(https://en.wikipedia.org/w/index.php?title=Natural_language_processing&action=edit&section=15 "Edit section: General tendencies and (possible) future directions")]

Based on long-standing trends in the field, it is possible to extrapolate future directions of NLP. As of 2020, three trends among the topics of the long-standing series of CoNLL Shared Tasks can be observed:[[47]](#cite_note-47)

- Interest on increasingly abstract, "cognitive" aspects of natural language (1999–2001: shallow parsing, 2002–03: named entity recognition, 2006–09/2017–18: dependency syntax, 2004–05/2008–09 semantic role labelling, 2011–12 coreference, 2015–16: discourse parsing, 2019: semantic parsing).
- Increasing interest in multilinguality, and, potentially, multimodality (English since 1999; Spanish, Dutch since 2002; German since 2003; Bulgarian, Danish, Japanese, Portuguese, Slovenian, Swedish, Turkish since 2006; Basque, Catalan, Chinese, Greek, Hungarian, Italian, Turkish since 2007; Czech since 2009; Arabic since 2012; 2017: 40+ languages; 2018: 60+/100+ languages)
- Elimination of symbolic representations (rule-based over supervised towards weakly supervised methods, representation learning and end-to-end systems)

Most higher-level NLP applications involve aspects that emulate intelligent behaviour and apparent comprehension of natural language. More broadly speaking, the technical operationalization of increasingly advanced aspects of cognitive behaviour represents one of the developmental trajectories of NLP (see trends among CoNLL shared tasks above).

[Cognition](https://en.wikipedia.org/wiki/Cognition "Cognition") refers to "the mental action or process of acquiring knowledge and understanding through thought, experience, and the senses."[[48]](#cite_note-48) [Cognitive science](https://en.wikipedia.org/wiki/Cognitive_science "Cognitive science") is the interdisciplinary, scientific study of the mind and its processes.[[49]](#cite_note-49) [Cognitive linguistics](https://en.wikipedia.org/wiki/Cognitive_linguistics "Cognitive linguistics") is an interdisciplinary branch of linguistics, combining knowledge and research from both psychology and linguistics.[[50]](#cite_note-50) Especially during the age of [symbolic NLP](#Symbolic_NLP_\(1950s_–_early_1990s\)), the area of computational linguistics maintained strong ties with cognitive studies.

As an example, [George Lakoff](https://en.wikipedia.org/wiki/George_Lakoff "George Lakoff") offers a methodology to build natural language processing (NLP) algorithms through the perspective of cognitive science, along with the findings of cognitive linguistics,[[51]](#cite_note-51) with two defining aspects:

1. Apply the theory of [conceptual metaphor](https://en.wikipedia.org/wiki/Conceptual_metaphor "Conceptual metaphor"), explained by Lakoff as "the understanding of one idea, in terms of another" which provides an idea of the intent of the author.[[52]](#cite_note-52) For example, consider the English word _big_. When used in a comparison ("That is a big tree"), the author's intent is to imply that the tree is _physically large_ relative to other trees or the authors experience. When used metaphorically ("Tomorrow is a big day"), the author's intent to imply _importance_. The intent behind other usages, like in "She is a big person", will remain somewhat ambiguous to a person and a cognitive NLP algorithm alike without additional information.
2. Assign relative measures of meaning to a word, phrase, sentence or piece of text based on the information presented before and after the piece of text being analyzed, e.g., by means of a [probabilistic context-free grammar](https://en.wikipedia.org/wiki/Probabilistic_context-free_grammar "Probabilistic context-free grammar") (PCFG). The mathematical equation for such algorithms is presented in [US Patent 9269353](https://worldwide.espacenet.com/patent/search/family/055314712/publication/US9269353B1?q=pn%3DUS9269353) [Archived](https://web.archive.org/web/20240516102600/https://worldwide.espacenet.com/patent/search/family/055314712/publication/US9269353B1?q=pn=US9269353) 2024-05-16 at the [Wayback Machine](https://en.wikipedia.org/wiki/Wayback_Machine "Wayback Machine"):[[53]](#cite_note-53)

![{\displaystyle {RMM(token_{N})}={PMM(token_{N})}\times {\frac {1}{2d}}\left(\sum _{i=-d}^{d}{((PMM(token_{N})}\times {PF(token_{N-i},token_{N},token_{N+i}))_{i}}\right)}](https://wikimedia.org/api/rest_v1/media/math/render/svg/43ccadd794a4b84e20d1209997a463342e0dfbfe)

_Where_

**RMM** is the relative measure of meaning

**token** is any block of text, sentence, phrase or word

**N** is the number of tokens being analyzed

**PMM** is the probable measure of meaning based on a corpora

**d** is the non zero location of the token along the sequence of **N** tokens

**PF** is the probability function specific to a language

Ties with cognitive linguistics are part of the historical heritage of NLP, but they have been less frequently addressed since the statistical turn during the 1990s. Nevertheless, approaches to develop cognitive models towards technically operationalizable frameworks have been pursued in the context of various frameworks, e.g., of cognitive grammar,[[54]](#cite_note-54) functional grammar,[[55]](#cite_note-55) construction grammar,[[56]](#cite_note-56) computational psycholinguistics and cognitive neuroscience (e.g., [ACT-R](https://en.wikipedia.org/wiki/ACT-R "ACT-R")), however, with limited uptake in mainstream NLP (as measured by presence on major conferences[[57]](#cite_note-57) of the [ACL](https://en.wikipedia.org/wiki/Association_for_Computational_Linguistics "Association for Computational Linguistics")). More recently, ideas of cognitive NLP have been revived as an approach to achieve [explainability](https://en.wikipedia.org/wiki/Explainable_artificial_intelligence "Explainable artificial intelligence"), e.g., under the notion of "cognitive AI".[[58]](#cite_note-58) Likewise, ideas of cognitive NLP are inherent to neural models [multimodal](https://en.wikipedia.org/wiki/Multimodal_interaction "Multimodal interaction") NLP (although rarely made explicit)[[59]](#cite_note-59) and developments in [artificial intelligence](https://en.wikipedia.org/wiki/Artificial_intelligence "Artificial intelligence"), specifically tools and technologies using [large language model](https://en.wikipedia.org/wiki/Large_language_model "Large language model") approaches[[60]](#cite_note-60) and new directions in [artificial general intelligence](https://en.wikipedia.org/wiki/Artificial_general_intelligence "Artificial general intelligence") based on the [free energy principle](https://en.wikipedia.org/wiki/Free_energy_principle "Free energy principle")[[61]](#cite_note-61) by British neuroscientist and theoretician at University College London [Karl J. Friston](https://en.wikipedia.org/wiki/Karl_J._Friston "Karl J. Friston").

- _[1 the Road](https://en.wikipedia.org/wiki/1_the_Road "1 the Road")_
- [Artificial intelligence detection software](https://en.wikipedia.org/wiki/Artificial_intelligence_detection_software "Artificial intelligence detection software")
- [Automated essay scoring](https://en.wikipedia.org/wiki/Automated_essay_scoring "Automated essay scoring")
- [Biomedical text mining](https://en.wikipedia.org/wiki/Biomedical_text_mining "Biomedical text mining")
- [Compound term processing](https://en.wikipedia.org/wiki/Compound_term_processing "Compound term processing")
- [Computational linguistics](https://en.wikipedia.org/wiki/Computational_linguistics "Computational linguistics")
- [Computer-assisted reviewing](https://en.wikipedia.org/wiki/Computer-assisted_reviewing "Computer-assisted reviewing")
- [Controlled natural language](https://en.wikipedia.org/wiki/Controlled_natural_language "Controlled natural language")
- [Deep learning](https://en.wikipedia.org/wiki/Deep_learning "Deep learning")
- [Deep linguistic processing](https://en.wikipedia.org/wiki/Deep_linguistic_processing "Deep linguistic processing")
- [Distributional semantics](https://en.wikipedia.org/wiki/Distributional_semantics "Distributional semantics")
- [Foreign language reading aid](https://en.wikipedia.org/wiki/Foreign_language_reading_aid "Foreign language reading aid")
- [Foreign language writing aid](https://en.wikipedia.org/wiki/Foreign_language_writing_aid "Foreign language writing aid")
- [Information extraction](https://en.wikipedia.org/wiki/Information_extraction "Information extraction")
- [Information retrieval](https://en.wikipedia.org/wiki/Information_retrieval "Information retrieval")
- [Language and Communication Technologies](https://en.wikipedia.org/wiki/Language_and_Communication_Technologies "Language and Communication Technologies")
- [Language model](https://en.wikipedia.org/wiki/Language_model "Language model")
- [Language technology](https://en.wikipedia.org/wiki/Language_technology "Language technology")
- [Latent semantic indexing](https://en.wikipedia.org/wiki/Latent_semantic_indexing "Latent semantic indexing")
- [Multi-agent system](https://en.wikipedia.org/wiki/Multi-agent_system "Multi-agent system")
- [Native-language identification](https://en.wikipedia.org/wiki/Native-language_identification "Native-language identification")
- [Natural-language programming](https://en.wikipedia.org/wiki/Natural-language_programming "Natural-language programming")
- [Natural-language understanding](https://en.wikipedia.org/wiki/Natural-language_understanding "Natural-language understanding")
- [Natural-language search](https://en.wikipedia.org/wiki/Natural-language_user_interface "Natural-language user interface")
- [Outline of natural language processing](https://en.wikipedia.org/wiki/Outline_of_natural_language_processing "Outline of natural language processing")
- [Query expansion](https://en.wikipedia.org/wiki/Query_expansion "Query expansion")
- [Query understanding](https://en.wikipedia.org/wiki/Query_understanding "Query understanding")
- [Reification (linguistics)](https://en.wikipedia.org/wiki/Reification_\(linguistics\) "Reification (linguistics)")
- [Speech processing](https://en.wikipedia.org/wiki/Speech_processing "Speech processing")
- [Spoken dialogue systems](https://en.wikipedia.org/wiki/Spoken_dialogue_systems "Spoken dialogue systems")
- [Text-proofing](https://en.wikipedia.org/wiki/Text-proofing "Text-proofing")
- [Text simplification](https://en.wikipedia.org/wiki/Text_simplification "Text simplification")
- [Transformer (machine learning model)](https://en.wikipedia.org/wiki/Transformer_\(machine_learning_model\) "Transformer (machine learning model)")
- [Truecasing](https://en.wikipedia.org/wiki/Truecasing "Truecasing")
- [Question answering](https://en.wikipedia.org/wiki/Question_answering "Question answering")
- [Word2vec](https://en.wikipedia.org/wiki/Word2vec "Word2vec")

1. **[^](#cite_ref-nlpintro_1-0 "Jump up")** Eisenstein, Jacob (October 1, 2019). [_Introduction to Natural Language Processing_](https://mitpress.mit.edu/9780262042840/introduction-to-natural-language-processing/). The MIT Press. p. 1. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-262-04284-0](https://en.wikipedia.org/wiki/Special:BookSources/978-0-262-04284-0 "Special:BookSources/978-0-262-04284-0").
2. **[^](#cite_ref-2 "Jump up")** ["NLP"](https://cs.stanford.edu/people/eroberts/courses/soco/projects/2004-05/nlp/overview_history.html).
3. **[^](#cite_ref-3 "Jump up")** Hutchins, J. (2005). ["The history of machine translation in a nutshell"](https://web.archive.org/web/20190713103044/http://www.hutchinsweb.me.uk/Nutshell-2005.pdf) (PDF). Archived from [the original](http://www.hutchinsweb.me.uk/Nutshell-2005.pdf) (PDF) on 2019-07-13. Retrieved 2019-02-04.[_[self-published source](https://en.wikipedia.org/wiki/Wikipedia:Verifiability#Self-published_sources "Wikipedia:Verifiability")_]
4. **[^](#cite_ref-4 "Jump up")** "ALPAC: the (in)famous report", John Hutchins, MT News International, no. 14, June 1996, pp. 9–12.
5. **[^](#cite_ref-5 "Jump up")** [Crevier 1993](#CITEREFCrevier1993), pp. 146–148, see also [Buchanan 2005](#CITEREFBuchanan2005), p. 56: "Early programs were necessarily limited in scope by the size and speed of memory"
6. **[^](#cite_ref-6 "Jump up")** [Koskenniemi, Kimmo](https://en.wikipedia.org/wiki/Kimmo_Koskenniemi "Kimmo Koskenniemi") (1983), [_Two-level morphology: A general computational model of word-form recognition and production_](https://web.archive.org/web/20181221032913/http://www.ling.helsinki.fi/~koskenni/doc/Two-LevelMorphology.pdf) (PDF), Department of General Linguistics, [University of Helsinki](https://en.wikipedia.org/wiki/University_of_Helsinki "University of Helsinki"), archived from [the original](http://www.ling.helsinki.fi/~koskenni/doc/Two-LevelMorphology.pdf) (PDF) on 2018-12-21, retrieved 2020-08-20
7. **[^](#cite_ref-7 "Jump up")** Joshi, A. K., & Weinstein, S. (1981, August). [Control of Inference: Role of Some Aspects of Discourse Structure-Centering](https://www.ijcai.org/Proceedings/81-1/Papers/071.pdf). In _IJCAI_ (pp. 385–387).
8. **[^](#cite_ref-8 "Jump up")** Guida, G.; Mauri, G. (July 1986). "Evaluation of natural language processing systems: Issues and approaches". _Proceedings of the IEEE_. **74** (7): 1026–1035. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1109/PROC.1986.13580](https://doi.org/10.1109%2FPROC.1986.13580). [ISSN](https://en.wikipedia.org/wiki/ISSN_\(identifier\) "ISSN (identifier)") [1558-2256](https://search.worldcat.org/issn/1558-2256). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [30688575](https://api.semanticscholar.org/CorpusID:30688575).
9. **[^](#cite_ref-9 "Jump up")** Chomskyan linguistics encourages the investigation of "[corner cases](https://en.wikipedia.org/wiki/Corner_case "Corner case")" that stress the limits of its theoretical models (comparable to [pathological](https://en.wikipedia.org/wiki/Pathological_\(mathematics\) "Pathological (mathematics)") phenomena in mathematics), typically created using [thought experiments](https://en.wikipedia.org/wiki/Thought_experiment "Thought experiment"), rather than the systematic investigation of typical phenomena that occur in real-world data, as is the case in [corpus linguistics](https://en.wikipedia.org/wiki/Corpus_linguistics "Corpus linguistics"). The creation and use of such [corpora](https://en.wikipedia.org/wiki/Text_corpus "Text corpus") of real-world data is a fundamental part of machine-learning algorithms for natural language processing. In addition, theoretical underpinnings of Chomskyan linguistics such as the so-called "[poverty of the stimulus](https://en.wikipedia.org/wiki/Poverty_of_the_stimulus "Poverty of the stimulus")" argument entail that general learning algorithms, as are typically used in machine learning, cannot be successful in language processing. As a result, the Chomskyan paradigm discouraged the application of such models to language processing.
10. **[^](#cite_ref-10 "Jump up")** Bengio, Yoshua; Ducharme, Réjean; Vincent, Pascal; Janvin, Christian (March 1, 2003). ["A neural probabilistic language model"](https://dl.acm.org/doi/10.5555/944919.944966). _The Journal of Machine Learning Research_. **3**: 1137–1155 – via ACM Digital Library.
11. **[^](#cite_ref-11 "Jump up")** Mikolov, Tomáš; Karafiát, Martin; Burget, Lukáš; Černocký, Jan; Khudanpur, Sanjeev (26 September 2010). ["Recurrent neural network based language model"](https://gwern.net/doc/ai/nn/rnn/2010-mikolov.pdf) (PDF). _Interspeech 2010_. pp. 1045–1048. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.21437/Interspeech.2010-343](https://doi.org/10.21437%2FInterspeech.2010-343). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [17048224](https://api.semanticscholar.org/CorpusID:17048224).
12. **[^](#cite_ref-goldberg:nnlp17_12-0 "Jump up")** Goldberg, Yoav (2016). "A Primer on Neural Network Models for Natural Language Processing". _Journal of Artificial Intelligence Research_. **57**: 345–420. [arXiv](https://en.wikipedia.org/wiki/ArXiv_\(identifier\) "ArXiv (identifier)"):[1807.10854](https://arxiv.org/abs/1807.10854). [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1613/jair.4992](https://doi.org/10.1613%2Fjair.4992). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [8273530](https://api.semanticscholar.org/CorpusID:8273530).
13. **[^](#cite_ref-goodfellow:book16_13-0 "Jump up")** Goodfellow, Ian; Bengio, Yoshua; Courville, Aaron (2016). [_Deep Learning_](http://www.deeplearningbook.org/). MIT Press.
14. **[^](#cite_ref-jozefowicz:lm16_14-0 "Jump up")** Jozefowicz, Rafal; Vinyals, Oriol; Schuster, Mike; Shazeer, Noam; Wu, Yonghui (2016). _Exploring the Limits of Language Modeling_. [arXiv](https://en.wikipedia.org/wiki/ArXiv_\(identifier\) "ArXiv (identifier)"):[1602.02410](https://arxiv.org/abs/1602.02410). [Bibcode](https://en.wikipedia.org/wiki/Bibcode_\(identifier\) "Bibcode (identifier)"):[2016arXiv160202410J](https://ui.adsabs.harvard.edu/abs/2016arXiv160202410J).
15. **[^](#cite_ref-choe:emnlp16_15-0 "Jump up")** Choe, Do Kook; Charniak, Eugene. ["Parsing as Language Modeling"](https://web.archive.org/web/20181023034804/https://aclanthology.coli.uni-saarland.de/papers/D16-1257/d16-1257). _Emnlp 2016_. Archived from [the original](https://aclanthology.coli.uni-saarland.de/papers/D16-1257/d16-1257) on 2018-10-23. Retrieved 2018-10-22.
16. **[^](#cite_ref-vinyals:nips15_16-0 "Jump up")** Vinyals, Oriol; et al. (2014). ["Grammar as a Foreign Language"](https://papers.nips.cc/paper/5635-grammar-as-a-foreign-language.pdf) (PDF). _Nips2015_. [arXiv](https://en.wikipedia.org/wiki/ArXiv_\(identifier\) "ArXiv (identifier)"):[1412.7449](https://arxiv.org/abs/1412.7449). [Bibcode](https://en.wikipedia.org/wiki/Bibcode_\(identifier\) "Bibcode (identifier)"):[2014arXiv1412.7449V](https://ui.adsabs.harvard.edu/abs/2014arXiv1412.7449V).
17. **[^](#cite_ref-17 "Jump up")** Turchin, Alexander; Florez Builes, Luisa F. (2021-03-19). ["Using Natural Language Processing to Measure and Improve Quality of Diabetes Care: A Systematic Review"](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC8120048). _Journal of Diabetes Science and Technology_. **15** (3): 553–560. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1177/19322968211000831](https://doi.org/10.1177%2F19322968211000831). [ISSN](https://en.wikipedia.org/wiki/ISSN_\(identifier\) "ISSN (identifier)") [1932-2968](https://search.worldcat.org/issn/1932-2968). [PMC](https://en.wikipedia.org/wiki/PMC_\(identifier\) "PMC (identifier)") [8120048](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC8120048). [PMID](https://en.wikipedia.org/wiki/PMID_\(identifier\) "PMID (identifier)") [33736486](https://pubmed.ncbi.nlm.nih.gov/33736486).
18. **[^](#cite_ref-18 "Jump up")** Lee, Jennifer; Yang, Samuel; Holland-Hall, Cynthia; Sezgin, Emre; Gill, Manjot; Linwood, Simon; Huang, Yungui; Hoffman, Jeffrey (2022-06-10). ["Prevalence of Sensitive Terms in Clinical Notes Using Natural Language Processing Techniques: Observational Study"](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9233261). _JMIR Medical Informatics_. **10** (6) e38482. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.2196/38482](https://doi.org/10.2196%2F38482). [ISSN](https://en.wikipedia.org/wiki/ISSN_\(identifier\) "ISSN (identifier)") [2291-9694](https://search.worldcat.org/issn/2291-9694). [PMC](https://en.wikipedia.org/wiki/PMC_\(identifier\) "PMC (identifier)") [9233261](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9233261). [PMID](https://en.wikipedia.org/wiki/PMID_\(identifier\) "PMID (identifier)") [35687381](https://pubmed.ncbi.nlm.nih.gov/35687381).
19. **[^](#cite_ref-winograd:shrdlu71_19-0 "Jump up")** Winograd, Terry (1971). [_Procedures as a Representation for Data in a Computer Program for Understanding Natural Language_](http://hci.stanford.edu/winograd/shrdlu/) (Thesis).
20. **[^](#cite_ref-schank77_20-0 "Jump up")** Schank, Roger C.; Abelson, Robert P. (1977). _Scripts, Plans, Goals, and Understanding: An Inquiry Into Human Knowledge Structures_. Hillsdale: Erlbaum. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [0-470-99033-3](https://en.wikipedia.org/wiki/Special:BookSources/0-470-99033-3 "Special:BookSources/0-470-99033-3").
21. **[^](#cite_ref-johnson:eacl:ilcl09_21-0 "Jump up")** [Mark Johnson. How the statistical revolution changes (computational) linguistics.](http://www.aclweb.org/anthology/W09-0103) Proceedings of the EACL 2009 Workshop on the Interaction between Linguistics and Computational Linguistics.
22. **[^](#cite_ref-resnik:langlog11_22-0 "Jump up")** [Philip Resnik. Four revolutions.](http://languagelog.ldc.upenn.edu/nll/?p=2946) Language Log, February 5, 2011.
23. **[^](#cite_ref-23 "Jump up")** Socher, Richard. ["Deep Learning For NLP-ACL 2012 Tutorial"](https://web.archive.org/web/20210414054126/https://www.socher.org/index.php/Main/DeepLearningForNLP-ACL2012Tutorial). _www.socher.org_. Archived from [the original](https://www.socher.org/index.php/Main/DeepLearningForNLP-ACL2012Tutorial) on 2021-04-14. Retrieved 2020-08-17. This was an early Deep Learning tutorial at the ACL 2012 and met with both interest and (at the time) skepticism by most participants. Until then, neural learning was basically rejected because of its lack of statistical interpretability. Until 2015, deep learning had evolved into the major framework of NLP. [Link is broken, try [http://web.stanford.edu/class/cs224n/](http://web.stanford.edu/class/cs224n/)]
24. **[^](#cite_ref-24 "Jump up")** Segev, Elad (2022). [_Semantic Network Analysis in Social Sciences_](https://www.routledge.com/Semantic-Network-Analysis-in-Social-Sciences/Segev/p/book/9780367636524). London: Routledge. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-367-63652-4](https://en.wikipedia.org/wiki/Special:BookSources/978-0-367-63652-4 "Special:BookSources/978-0-367-63652-4"). [Archived](https://web.archive.org/web/20211205140726/https://www.routledge.com/Semantic-Network-Analysis-in-Social-Sciences/Segev/p/book/9780367636524) from the original on 5 December 2021. Retrieved 5 December 2021.
25. **[^](#cite_ref-25 "Jump up")** Yi, Chucai; [Tian, Yingli](https://en.wikipedia.org/wiki/Yingli_Tian "Yingli Tian") (2012), "Assistive Text Reading from Complex Background for Blind Persons", _Camera-Based Document Analysis and Recognition_, Lecture Notes in Computer Science, vol. 7139, Springer Berlin Heidelberg, pp. 15–28, [CiteSeerX](https://en.wikipedia.org/wiki/CiteSeerX_\(identifier\) "CiteSeerX (identifier)") [10.1.1.668.869](https://citeseerx.ist.psu.edu/viewdoc/summary?doi=10.1.1.668.869), [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1007/978-3-642-29364-1_2](https://doi.org/10.1007%2F978-3-642-29364-1_2), [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-3-642-29363-4](https://en.wikipedia.org/wiki/Special:BookSources/978-3-642-29363-4 "Special:BookSources/978-3-642-29363-4")
26. ^ [Jump up to: _**a**_](#cite_ref-:0_26-0) [_**b**_](#cite_ref-:0_26-1) ["Natural Language Processing (NLP) - A Complete Guide"](https://www.deeplearning.ai/resources/natural-language-processing/). _www.deeplearning.ai_. 2023-01-11. Retrieved 2024-05-05.
27. **[^](#cite_ref-27 "Jump up")** ["What is Natural Language Processing? Intro to NLP in Machine Learning"](https://www.gyansetu.in/what-is-natural-language-processing/). _GyanSetu!_. 2020-12-06. Retrieved 2021-01-09.
28. **[^](#cite_ref-28 "Jump up")** Kishorjit, N.; Vidya, Raj RK.; Nirmal, Y.; Sivaji, B. (2012). ["Manipuri Morpheme Identification"](http://aclweb.org/anthology//W/W12/W12-5008.pdf) (PDF). _Proceedings of the 3rd Workshop on South and Southeast Asian Natural Language Processing (SANLP)_. COLING 2012, Mumbai, December 2012: 95–108.`{{[cite journal](https://en.wikipedia.org/wiki/Template:Cite_journal "Template:Cite journal")}}`: CS1 maint: location ([link](https://en.wikipedia.org/wiki/Category:CS1_maint:_location "Category:CS1 maint: location"))
29. **[^](#cite_ref-29 "Jump up")** Klein, Dan; Manning, Christopher D. (2002). ["Natural language grammar induction using a constituent-context model"](http://papers.nips.cc/paper/1945-natural-language-grammar-induction-using-a-constituent-context-model.pdf) (PDF). _Advances in Neural Information Processing Systems_.
30. **[^](#cite_ref-30 "Jump up")** Kariampuzha, William; Alyea, Gioconda; Qu, Sue; Sanjak, Jaleal; Mathé, Ewy; Sid, Eric; Chatelaine, Haley; Yadaw, Arjun; Xu, Yanji; Zhu, Qian (2023). ["Precision information extraction for rare disease epidemiology at scale"](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9972634). _Journal of Translational Medicine_. **21** (1): 157. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1186/s12967-023-04011-y](https://doi.org/10.1186%2Fs12967-023-04011-y). [PMC](https://en.wikipedia.org/wiki/PMC_\(identifier\) "PMC (identifier)") [9972634](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC9972634). [PMID](https://en.wikipedia.org/wiki/PMID_\(identifier\) "PMID (identifier)") [36855134](https://pubmed.ncbi.nlm.nih.gov/36855134).
31. **[^](#cite_ref-rte:11_31-0 "Jump up")** PASCAL Recognizing Textual Entailment Challenge (RTE-7) [https://tac.nist.gov//2011/RTE/](https://tac.nist.gov//2011/RTE/)
32. **[^](#cite_ref-32 "Jump up")** Lippi, Marco; Torroni, Paolo (2016-04-20). ["Argumentation Mining: State of the Art and Emerging Trends"](https://dl.acm.org/doi/10.1145/2850417). _ACM Transactions on Internet Technology_. **16** (2): 1–25. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1145/2850417](https://doi.org/10.1145%2F2850417). [hdl](https://en.wikipedia.org/wiki/Hdl_\(identifier\) "Hdl (identifier)"):[11585/523460](https://hdl.handle.net/11585%2F523460). [ISSN](https://en.wikipedia.org/wiki/ISSN_\(identifier\) "ISSN (identifier)") [1533-5399](https://search.worldcat.org/issn/1533-5399). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [9561587](https://api.semanticscholar.org/CorpusID:9561587).
33. **[^](#cite_ref-33 "Jump up")** ["Argument Mining – IJCAI2016 Tutorial"](https://web.archive.org/web/20210418083659/https://www.i3s.unice.fr/~villata/tutorialIJCAI2016.html). _www.i3s.unice.fr_. Archived from [the original](https://www.i3s.unice.fr/~villata/tutorialIJCAI2016.html) on 2021-04-18. Retrieved 2021-03-09.
34. **[^](#cite_ref-34 "Jump up")** ["NLP Approaches to Computational Argumentation – ACL 2016, Berlin"](http://acl2016tutorial.arg.tech/). Retrieved 2021-03-09.
35. **[^](#cite_ref-35 "Jump up")** Administration. ["Centre for Language Technology (CLT)"](https://www.mq.edu.au/research/research-centres-groups-and-facilities/innovative-technologies/centres/centre-for-language-technology-clt). _Macquarie University_. Retrieved 2021-01-11.
36. **[^](#cite_ref-36 "Jump up")** ["Shared Task: Grammatical Error Correction"](https://www.comp.nus.edu.sg/~nlp/conll13st.html). _www.comp.nus.edu.sg_. Retrieved 2021-01-11.
37. **[^](#cite_ref-37 "Jump up")** ["Shared Task: Grammatical Error Correction"](https://www.comp.nus.edu.sg/~nlp/conll14st.html). _www.comp.nus.edu.sg_. Retrieved 2021-01-11.
38. **[^](#cite_ref-38 "Jump up")** Duan, Yucong; Cruz, Christophe (2011). ["Formalizing Semantic of Natural Language through Conceptualization from Existence"](https://web.archive.org/web/20111009135952/http://www.ijimt.org/abstract/100-E00187.htm). _International Journal of Innovation, Management and Technology_. **2** (1): 37–42. Archived from [the original](http://www.ijimt.org/abstract/100-E00187.htm) on 2011-10-09.
39. **[^](#cite_ref-39 "Jump up")** ["U B U W E B :: Racter"](http://www.ubu.com/historical/racter/index.html). _www.ubu.com_. Retrieved 2020-08-17.
40. **[^](#cite_ref-40 "Jump up")** Writer, Beta (2019). _Lithium-Ion Batteries_. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1007/978-3-030-16800-1](https://doi.org/10.1007%2F978-3-030-16800-1). [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-3-030-16799-8](https://en.wikipedia.org/wiki/Special:BookSources/978-3-030-16799-8 "Special:BookSources/978-3-030-16799-8"). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [155818532](https://api.semanticscholar.org/CorpusID:155818532).
41. **[^](#cite_ref-41 "Jump up")** ["Document Understanding AI on Google Cloud (Cloud Next '19) – YouTube"](https://ghostarchive.org/varchive/youtube/20211030/7dtl650D0y0). _www.youtube.com_. 11 April 2019. Archived from [the original](https://www.youtube.com/watch?v=7dtl650D0y0) on 2021-10-30. Retrieved 2021-01-11.
42. **[^](#cite_ref-42 "Jump up")** Robertson, Adi (2022-04-06). ["OpenAI's DALL-E AI image generator can now edit pictures, too"](https://www.theverge.com/2022/4/6/23012123/openai-clip-dalle-2-ai-text-to-image-generator-testing). _The Verge_. Retrieved 2022-06-07.
43. **[^](#cite_ref-43 "Jump up")** ["The Stanford Natural Language Processing Group"](https://nlp.stanford.edu/projects/text2scene.shtml). _nlp.stanford.edu_. Retrieved 2022-06-07.
44. **[^](#cite_ref-44 "Jump up")** Coyne, Bob; Sproat, Richard (2001-08-01). ["WordsEye"](https://doi.org/10.1145/383259.383316). _Proceedings of the 28th annual conference on Computer graphics and interactive techniques_. SIGGRAPH '01. New York, NY, USA: Association for Computing Machinery. pp. 487–496. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1145/383259.383316](https://doi.org/10.1145%2F383259.383316). [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-1-58113-374-5](https://en.wikipedia.org/wiki/Special:BookSources/978-1-58113-374-5 "Special:BookSources/978-1-58113-374-5"). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [3842372](https://api.semanticscholar.org/CorpusID:3842372).
45. **[^](#cite_ref-45 "Jump up")** ["Google announces AI advances in text-to-video, language translation, more"](https://venturebeat.com/ai/google-announces-ai-advances-in-text-to-video-language-translation-more/). _VentureBeat_. 2022-11-02. Retrieved 2022-11-09.
46. **[^](#cite_ref-46 "Jump up")** Vincent, James (2022-09-29). ["Meta's new text-to-video AI generator is like DALL-E for video"](https://www.theverge.com/2022/9/29/23378210/meta-text-to-video-ai-generation-make-a-video-model-dall-e). _The Verge_. Retrieved 2022-11-09.
47. **[^](#cite_ref-47 "Jump up")** ["Previous shared tasks | CoNLL"](https://www.conll.org/previous-tasks). _www.conll.org_. Retrieved 2021-01-11.
48. **[^](#cite_ref-48 "Jump up")** ["Cognition"](https://web.archive.org/web/20200715113427/https://www.lexico.com/definition/cognition). _Lexico_. [Oxford University Press](https://en.wikipedia.org/wiki/Oxford_University_Press "Oxford University Press") and [Dictionary.com](https://en.wikipedia.org/wiki/Dictionary.com "Dictionary.com"). Archived from [the original](https://www.lexico.com/definition/cognition) on July 15, 2020. Retrieved 6 May 2020.
49. **[^](#cite_ref-49 "Jump up")** ["Ask the Cognitive Scientist"](http://www.aft.org/newspubs/periodicals/ae/summer2002/willingham.cfm). _American Federation of Teachers_. 8 August 2014. Cognitive science is an interdisciplinary field of researchers from Linguistics, psychology, neuroscience, philosophy, computer science, and anthropology that seek to understand the mind.
50. **[^](#cite_ref-50 "Jump up")** Robinson, Peter (2008). _Handbook of Cognitive Linguistics and Second Language Acquisition_. Routledge. pp. 3–8. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-805-85352-0](https://en.wikipedia.org/wiki/Special:BookSources/978-0-805-85352-0 "Special:BookSources/978-0-805-85352-0").
51. **[^](#cite_ref-51 "Jump up")** Lakoff, George (1999). _Philosophy in the Flesh: The Embodied Mind and Its Challenge to Western Philosophy; Appendix: The Neural Theory of Language Paradigm_. New York Basic Books. pp. 569–583. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-465-05674-3](https://en.wikipedia.org/wiki/Special:BookSources/978-0-465-05674-3 "Special:BookSources/978-0-465-05674-3").
52. **[^](#cite_ref-52 "Jump up")** Strauss, Claudia (1999). _A Cognitive Theory of Cultural Meaning_. Cambridge University Press. pp. 156–164. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-521-59541-4](https://en.wikipedia.org/wiki/Special:BookSources/978-0-521-59541-4 "Special:BookSources/978-0-521-59541-4").
53. **[^](#cite_ref-53 "Jump up")** [US patent 9269353](https://worldwide.espacenet.com/textdoc?DB=EPODOC&IDX=US9269353)
54. **[^](#cite_ref-54 "Jump up")** ["Universal Conceptual Cognitive Annotation (UCCA)"](https://universalconceptualcognitiveannotation.github.io/). _Universal Conceptual Cognitive Annotation (UCCA)_. Retrieved 2021-01-11.
55. **[^](#cite_ref-55 "Jump up")** Rodríguez, F. C., & Mairal-Usón, R. (2016). [Building an RRG computational grammar](https://www.redalyc.org/pdf/1345/134549291020.pdf). _Onomazein_, (34), 86–117.
56. **[^](#cite_ref-56 "Jump up")** ["Fluid Construction Grammar – A fully operational processing system for construction grammars"](https://www.fcg-net.org/). Retrieved 2021-01-11.
57. **[^](#cite_ref-57 "Jump up")** ["ACL Member Portal | The Association for Computational Linguistics Member Portal"](https://www.aclweb.org/portal/). _www.aclweb.org_. Retrieved 2021-01-11.
58. **[^](#cite_ref-58 "Jump up")** ["Chunks and Rules"](https://www.w3.org/Data/demos/chunks/chunks.html). _W3C_. Retrieved 2021-01-11.
59. **[^](#cite_ref-59 "Jump up")** Socher, Richard; Karpathy, Andrej; Le, Quoc V.; Manning, Christopher D.; Ng, Andrew Y. (2014). ["Grounded Compositional Semantics for Finding and Describing Images with Sentences"](https://doi.org/10.1162%2Ftacl_a_00177). _Transactions of the Association for Computational Linguistics_. **2**: 207–218. [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1162/tacl_a_00177](https://doi.org/10.1162%2Ftacl_a_00177). [S2CID](https://en.wikipedia.org/wiki/S2CID_\(identifier\) "S2CID (identifier)") [2317858](https://api.semanticscholar.org/CorpusID:2317858).
60. **[^](#cite_ref-60 "Jump up")** Dasgupta, Ishita; Lampinen, Andrew K.; Chan, Stephanie C. Y.; Creswell, Antonia; Kumaran, Dharshan; McClelland, James L.; Hill, Felix (2022). "Language models show human-like content effects on reasoning, Dasgupta, Lampinen et al". [arXiv](https://en.wikipedia.org/wiki/ArXiv_\(identifier\) "ArXiv (identifier)"):[2207.07051](https://arxiv.org/abs/2207.07051) [[cs.CL](https://arxiv.org/archive/cs.CL)].
61. **[^](#cite_ref-61 "Jump up")** Friston, Karl J. (2022). _Active Inference: The Free Energy Principle in Mind, Brain, and Behavior; Chapter 4 The Generative Models of Active Inference_. The MIT Press. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-262-36997-8](https://en.wikipedia.org/wiki/Special:BookSources/978-0-262-36997-8 "Special:BookSources/978-0-262-36997-8").

- Bates, M (1995). ["Models of natural language understanding"](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC40721). _Proceedings of the National Academy of Sciences of the United States of America_. **92** (22): 9977–9982. [Bibcode](https://en.wikipedia.org/wiki/Bibcode_\(identifier\) "Bibcode (identifier)"):[1995PNAS...92.9977B](https://ui.adsabs.harvard.edu/abs/1995PNAS...92.9977B). [doi](https://en.wikipedia.org/wiki/Doi_\(identifier\) "Doi (identifier)"):[10.1073/pnas.92.22.9977](https://doi.org/10.1073%2Fpnas.92.22.9977). [PMC](https://en.wikipedia.org/wiki/PMC_\(identifier\) "PMC (identifier)") [40721](https://www.ncbi.nlm.nih.gov/pmc/articles/PMC40721). [PMID](https://en.wikipedia.org/wiki/PMID_\(identifier\) "PMID (identifier)") [7479812](https://pubmed.ncbi.nlm.nih.gov/7479812).
- Steven Bird, Ewan Klein, and Edward Loper (2009). _Natural Language Processing with Python_. O'Reilly Media. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-596-51649-9](https://en.wikipedia.org/wiki/Special:BookSources/978-0-596-51649-9 "Special:BookSources/978-0-596-51649-9").
- [Kenna Hughes-Castleberry](https://en.wikipedia.org/w/index.php?title=Kenna_Hughes-Castleberry&action=edit&redlink=1 "Kenna Hughes-Castleberry (page does not exist)"), "A Murder Mystery Puzzle: The literary puzzle _[Cain's Jawbone](https://en.wikipedia.org/wiki/Cain%27s_Jawbone "Cain's Jawbone")_, which has stumped humans for decades, reveals the limitations of natural-language-processing algorithms", _[Scientific American](https://en.wikipedia.org/wiki/Scientific_American "Scientific American")_, vol. 329, no. 4 (November 2023), pp. 81–82. "This murder mystery competition has revealed that although NLP ([natural-language processing](https://en.wikipedia.org/wiki/Natural-language_processing "Natural-language processing")) models are capable of incredible feats, their abilities are very much limited by the amount of [context](https://en.wikipedia.org/wiki/Context_\(linguistics\) "Context (linguistics)") they receive. This [...] could cause [difficulties] for researchers who hope to use them to do things such as analyze [ancient languages](https://en.wikipedia.org/wiki/Ancient_language "Ancient language"). In some cases, there are few historical records on long-gone [civilizations](https://en.wikipedia.org/wiki/Civilization "Civilization") to serve as [training data](https://en.wikipedia.org/wiki/Training_data "Training data") for such a purpose." (p. 82.)
- Daniel Jurafsky and James H. Martin (2008). _Speech and Language Processing_, 2nd edition. Pearson Prentice Hall. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-13-187321-6](https://en.wikipedia.org/wiki/Special:BookSources/978-0-13-187321-6 "Special:BookSources/978-0-13-187321-6").
- Mohamed Zakaria Kurdi (2016). _Natural Language Processing and Computational Linguistics: speech, morphology, and syntax_, Volume 1. ISTE-Wiley. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-1848218482](https://en.wikipedia.org/wiki/Special:BookSources/978-1848218482 "Special:BookSources/978-1848218482").
- Mohamed Zakaria Kurdi (2017). _Natural Language Processing and Computational Linguistics: semantics, discourse, and applications_, Volume 2. ISTE-Wiley. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-1848219212](https://en.wikipedia.org/wiki/Special:BookSources/978-1848219212 "Special:BookSources/978-1848219212").
- Christopher D. Manning, Prabhakar Raghavan, and Hinrich Schütze (2008). _Introduction to Information Retrieval_. Cambridge University Press. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-521-86571-5](https://en.wikipedia.org/wiki/Special:BookSources/978-0-521-86571-5 "Special:BookSources/978-0-521-86571-5"). [Official html and pdf versions available without charge.](http://nlp.stanford.edu/IR-book/)
- Christopher D. Manning and Hinrich Schütze (1999). _Foundations of Statistical Natural Language Processing_. The MIT Press. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-262-13360-9](https://en.wikipedia.org/wiki/Special:BookSources/978-0-262-13360-9 "Special:BookSources/978-0-262-13360-9").
- David M. W. Powers and Christopher C. R. Turk (1989). _Machine Learning of Natural Language_. Springer-Verlag. [ISBN](https://en.wikipedia.org/wiki/ISBN_\(identifier\) "ISBN (identifier)") [978-0-387-19557-5](https://en.wikipedia.org/wiki/Special:BookSources/978-0-387-19557-5 "Special:BookSources/978-0-387-19557-5").

- [![](https://upload.wikimedia.org/wikipedia/en/thumb/4/4a/Commons-logo.svg/20px-Commons-logo.svg.png)](https://en.wikipedia.org/wiki/File:Commons-logo.svg) Media related to [Natural language processing](https://commons.wikimedia.org/wiki/Category:Natural_language_processing "commons:Category:Natural language processing") at Wikimedia Commons