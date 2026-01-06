## Structured Markdown

# AI-assisted JSON Schema Creation and Mapping

## Pre-print Information

  - **arXiv ID**: `2508.05192v1 [cs.SE]`
  - **Date**: `7 Aug 2025`
  - **Status**: `Author pre-print. Publication accepted for MODELS'25.`

### Authors and Affiliations

1.  **Felix Neubauer** (1st) and **Benjamin Uekermann** (3rd)
      - Institute for Parallel and Distributed Systems, University of Stuttgart, Stuttgart, Germany
      - Emails: `Felix.Neubauer@ipvs.uni-stuttgart.de`, `Benjamin. Uekermann@ipvs.uni-stuttgart.de`
2.  **Jürgen Pleiss** (2nd)
      - Institute of Biochemistry and Technical Biochemistry, University of Stuttgart, Stuttgart, Germany
      - Email: `Juergen.Pleiss@itb.uni-stuttgart.de`

-----

## Abstract

Model-Driven Engineering (MDE) places models at the core of system and data engineering processes [285]. In the context of research data, these models are typically expressed as **schemas** that define the structure and semantics of datasets [286]. However, many domains still lack standardized models, and creating them remains a significant barrier, especially for non-experts [287].

The authors present a **hybrid approach** that combines **large language models (LLMs)** with deterministic techniques to enable **JSON Schema creation, modification, and schema mapping** based on natural language inputs by the user [288].

These capabilities are integrated into the open-source tool **MetaConfigurator** [289, 307], which already provides visual model editing, validation, code generation, and form generation from models [289, 308].

For data integration, the tool generates **schema mappings** from heterogeneous data sources (`JSON`, `CSV`, `XML`, and `YAML`) using LLMs, while ensuring scalability and reliability through deterministic execution of generated mapping rules [290].

The applicability of this work is demonstrated in an application example in the field of **chemistry** [291]. By combining natural language interaction with deterministic safeguards, this work significantly lowers the barrier to structured data modeling and data integration for non-experts [292].

-----

## Index Terms

`rdm`, `research data management`, `mde`, `model driven engineering`, `json`, `yaml`, `configuration`, `schema`, `data`, `model`, `editor`, `gui`, `tool`, `ai`, `llm`, `mapping`, `matching` [293].

-----

## I. INTRODUCTION

Model-Driven Engineering (MDE) emphasizes the central role of models in the design, development, and operation of systems [1, 295]. In the context of research data management [2]–[4, 296], models are typically expressed as schemas, which define the structure and semantics of datasets [296]. These models enable key functionalities:

  - Instance validation, transformation, code generation, and visualization [297].
  - Ensuring data quality, interoperability, and reusability [297].

**Problem:** Many scientific domains still lack standardized data models [298]. For example, in chemistry, data is often stored in **Electronic Lab Notebooks (ELNs)** or spreadsheets, where much of the meaning is implicit [299]. Furthermore, `CSV` documents are limited to flat, table-like structures and cannot encode relationships, constraints, or domain-specific rules [300]. This lack of structure restricts interoperability, reproducibility, and the application of downstream methods such as knowledge graph construction or automated schema-driven validation [301, 303].

**Solution with LLMs:** Recent advances in LLMs offer a promising opportunity to bridge this gap by assisting users, particularly non-experts, in the creation and refinement of data models [304]. LLMs can translate natural language descriptions into structured model representations, potentially democratizing model-driven techniques [305].

**Limitations of relying solely on LLMs** [306]:

1.  Lack of guaranteed model validity.
2.  Limited interpretability of plain-text outputs.
3.  Challenges in processing large or complex datasets.
4.  The need for users to craft effective prompts.

**The Proposed MDE-based Approach (MetaConfigurator)** [307]:

The presented approach addresses these issues by integrating LLMs with **deterministic safeguards**, including targeted pre- and post-processing, and the rule-based execution of AI-generated mappings, implemented in the open-source tool **MetaConfigurator** [5, 307] (`¹https://github.com/MetaConfigurator/meta-configurator`).

1.  **Schema Capabilities (Section III-A):** The tool is extended with capabilities for **schema creation, modification, and querying from natural language input** [309]. This integrates prompt engineering, context management, and post-processing safeguards with visual feedback and schema validation, aligning LLM-generated content with MDE principles of correctness and transparency [310].
2.  **Data Integration (Section III-B):** A model-driven approach is introduced for data integration [311]. Given heterogeneous data sources (`JSON` [6], `YAML` [7], `XML` [8], or `CSV`), LLMs generate **human-readable mapping rules** to transform the data into a target schema (model) [311].
      - These rules are executed **deterministically**, separating generation from execution and ensuring reliability and scalability, especially for large datasets [312].

**Demonstration:** Section IV showcases an example from the domain of chemistry, where existing unstructured Excel data are transformed into structured, interoperable, and AI-ready representations using the extended toolchain [313].

-----

## II. BACKGROUND AND RELATED WORK

### A. JSON and JSON Schema

  - **JSON** is a widespread data format for exchanging information with web services, storing semi-structured documents in NoSQL databases [9], and representing structured content in APIs and configuration files [319]. It is suitable for data interchange and machine learning pipelines due to its simplicity, human-readability, and wide support [320].
  - **JSON Schema** is the de-facto standard for describing the structure and constraints of JSON documents [10, 321]. It plays a crucial role in enabling validation, interoperability, and automation [321]. It can be understood as a modeling language, similar to class diagrams in MDE [322].

### B. JSON Schema Creation

JSON schema editors are tools for creating and editing schemas [324]. Examples include `MetaConfigurator` [5], `Adamant` [11], and `Liquid Studio JSON Schema Editor` (paid) [325]. All these editors require some level of JSON schema understanding [326].

  - **MetaConfigurator** is used in this work as a general-purpose schema editor and form generator that supports different data formats (e.g., `JSON`, `YAML`, `XML`) and editing methods (e.g., text editor, GUI editor, or schema diagram [12]) [327]. It was chosen for its:
      - Modular architecture [328].
      - Ability to generate source code in 17 programming languages and generate documentation [328].
      - Open source, free, and accessible as a web service [329].

Approaches also exist for automatically generating a schema based on an instance dataset, referred to as **schema inference** [13] or **schema discovery** [14] tasks [330, 331].

**Issues with pure LLM-based Schema Generation (e.g., ChatGPT)** [332]:

  - Lack proper schema validation and graphical visualization [332].
  - LLMs can be distracted by irrelevant contexts [15, 333].
  - Accuracy drops significantly for low-probability inputs, even in deterministic tasks [16, 334].
  - Reasoning abilities degrade as input length increases [17, 335].
  - *Related Work:* Mior [18] fine-tuned an LLM for schema-related tasks, outperforming the base model `Code LLama` [19] significantly [336].

*URL for paid editor:* `https://www.liquid-technologies.com/json-schema-editor, acc. 25/06/02` [342].

### C. Schema Matching and Mapping

Integrating data from various sources and formats is a significant challenge [20]–[22, 338]. This work studies the task of **schema mapping**: converting a JSON instance from one schema to another [339]. A mapping can be simple property-to-property correspondence or include more expressive logics [340].

**Existing Transformation Tools:**

  - `JSON to JSON transformation (Jolt)` (Java library) [341]. (`³https://github.com/bazaarvoice/jolt, acc. 25/06/02`)
  - `JSONata` (TypeScript library) [341]. (`⁴https://github.com/jsonata-js/jsonata, acc. 25/06/02`)
  - `jq` (JSON query language) [343]. (`⁵https://github.com/jqlang/jq, acc. 25/06/02`) - While powerful for concise querying, it is not primarily intended for large documents due to its in-memory processing model and limited streaming/parallel execution support [344].

**Mapping Creation:** Can be done manually or through automated approaches [23, 345].

  - **Schema matching** is the task of identifying corresponding elements between two schemas [346]. Rahm and Bernstein [24] compared automated schema matching approaches for relational databases [347].
  - Stanek and Killough [25] created a program to generate code for transforming a JSON document, defining and implementing mapping rules themselves [348].
  - Buss et al. [26] discuss using LLMs for automatic schema mapping in relational databases, noting that **human-in-the-loop** approaches are required due to the difficult nature of data integration [349, 350].

### D. Large Language Models (LLMs) and Prompt Engineering

  - **LLMs** are deep learning models trained on large text corpora to perform a wide range of natural language processing tasks [27]–[29, 352].
  - **Prompt engineering** is the task of optimizing the input prompt for an LLM to get the best results [353].

**Prompt Engineering Techniques** [30, 354]:

  - **Few-shot training** (providing input-output examples) improves performance on complex tasks but requires additional tokens and can lead to biases [355].
  - Assigning the LLM a **role or persona** can increase performance [31, 356].
  - Reframing complex long instructional prompts into multiple smaller **subtasks** has shown to improve results [32, 357].
  - Providing **detailed natural-language instructions** significantly improves task performance compared to relying on implicit task formulation [33]–[35, 358].

-----

## III. DESIGN AND IMPLEMENTATION

The system describes the design and implementation of AI-assisted schema creation (Section III-A) and schema matching (Section III-B) approaches [359, 360].

The system communicates with LLMs via a configurable endpoint following the **OpenAI API** [361] (`⁶https://platform.openai.com/docs, acc. 25/06/07`). Users select the desired model via application settings [362]. Prompts are programmatically constructed, and resulting textual outputs are automatically parsed and processed [363]. Users must supply their own credentials for API access [364, 365]. All presented interactions and evaluations were conducted using **`gpt-40-mini`**, chosen for its cost-efficiency and proficiency in JSON Schema tasks [366, 368].

### A. AI-assisted JSON Schema Creation

This functionality is enabled by integrating a new AI assistance view into MetaConfigurator's modular architecture [369, 370], providing a conversational, chat-like interface [371].

The **hybrid method** offers several advantages over pure LLM-based solutions [372]:

1.  **Prompt Construction and Context Management** [373]:
      - The tool handles prompt engineering, dynamically constructing a structured prompt that sets the LLM's role (e.g., "You are a JSON Schema expert") [374].
      - For modifications, only the relevant schema subset, based on the user's current selection, is included as context [375, 378] (to prevent inaccuracies/hallucinations on large schemas [377]).
2.  **Integrated Validation and Visualization** [376]: The generated or modified schema is immediately validated and visualized within MetaConfigurator's schema editor [376].
3.  **Automated Response Post-processing** [379]: The tool cleans the LLM response by removing formatting artifacts (e.g., code fences or language identifiers) [379].
4.  **Human-in-the-loop Editing** [380]: If the LLM output is incomplete or incorrect, the user can manually correct the schema before accepting or discarding the change [380, 381].

The chat-based interface is used for creation (Figure 1), modification (Figure 2), querying a schema, or creating/editing/querying document instances [384].

**Example of Schema Creation (Figure 1):**

  - **User Prompt (Top):** "Create a schema about MOF synthesis in chemistry. We have a list of synthesis experiments. Each experiment has a metal salt and a ligand and also creator, date, temperature, duration, product\_purity (boolean). Ligand and metal salt should be objects which each have name, mass, inchi as properties." [390]
  - **Resulting Visual Schema (Bottom):** Shows `root`, `experiments entry`, `metalSalt`, and `ligand` structures, including property types (e.g., `string`, `number`, `boolean`) and array types (`object[]`) [391]-[400].

**Example of Schema Modification (Figure 2):**

  - **User Prompt (Top):** "The metal salt and ligand both have the same structure, both are compounds. Instead of having two different definitions, create one compound class with their properties and then make metal\_salt and ligand both refer to this compound class." [402]
  - **Updated Visual Schema (Bottom):** Shows the consolidation into a `compound` class, with `metal_salt` and `ligand` referencing it [403]-[417].

### B. AI-assisted Schema Matching and Mapping

Schema mapping uses human-defined rules, traditional automated approaches [23, 25], and recent LLM-based methods [26, 386]. The new implemented approach is:

1.  LLM generates **schema mapping rules** [387].
2.  These rules are then executed **deterministically** [387].

MetaConfigurator supports `JSON`, `YAML`, and `XML`, and can import `CSV` (by conversion to JSON), making data from any of these formats transformable [388, 389].

The schema mapping language used is **JSONata** [341], which is highly expressive and supports complex features (e.g., numeric, path, boolean/comparison operators, sorting, grouping, aggregation, functions, and functional programming) [419].

**Mapping Process** [420]:

1.  The user provides an **input JSON document** and a **target JSON schema** [420].
2.  A **source schema** is inferred for the input JSON instance using the `jsonhero/schema-infer` library [421]. (`⁷https://github.com/triggerdotdev/schema-infer, acc. 25/06/05.`)
3.  The **LLM prompt** is constructed, including mapping instructions (role/tasks), a mapping example (input/expected output), and the user input [422]. Detailed JSONata specifications are included to improve the quality of generated expressions [423, 425] (`⁸https://doi.org/10.18419/DARUS-5157`).
4.  **Array and properties truncation** is performed on potentially large input JSON documents to shorten them while maintaining overall structure (target size 64KB, minimum $n_{min}=2$) [426, 427]. To mitigate information loss, the JSON schema inferred from the *complete* document is included in the prompt [429].
5.  A **post-processing algorithm** is executed on the LLM's response, removing artifacts like code fences (often with a `jsonata` hint) [431, 432].
6.  The suggested mapping is presented in an interactive code editor, automatically validated for **syntactical correctness** (not semantics) using the JSONata library [433, 434].
7.  The user can make changes and can only apply the mapping once it is recognized as valid (Human-in-the-loop) [435].

**Example of Schema Mapping (Figure 3):**

  - **Input JSON Document / Target Schema (Top):** Shows the source data structure (e.g., `mof_id`, `reagents`) and the desired target structure (`materialId`, `metal`, `linkers`) [448]-[463].
  - **Mapping Generated by LLM (Middle):** JSONata expression is generated: `"materialId": mof_id, "metal": reagents[role = "metal"].name, "linkers": reagents[role = "linker"].name` [464]-[467].
  - **Transformed JSON Document (Bottom):** Shows the structured output based on the generated mapping [468]-[471].

-----

## IV. APPLICATION EXAMPLE

Chemists tracked inputs and results of chemical synthesis experiments in an **Excel table**, which is not suitable for machine learning algorithms [438, 439].

1.  **Conversion and Inference:** MetaConfigurator converts the Excel table into a JSON document and automatically infers the corresponding schema [439].
2.  **Schema Refinement:** The inferred schema is insufficient. For instance, `product_purity` values (`yes` and `no`) were inferred as `string` instead of `boolean`, and the `ligand` and `metal_salt` properties were listed flatly instead of being extracted into separate classes [440, 441].
3.  **AI-Assisted Improvement:** Using the chat-based interface (as per Figure 1-2), an improved schema is created [442].
4.  **Data Transformation:** The automated schema mapping functionality is used to transform the JSON document to satisfy the new schema [443]. This successfully maps `yes` and `no` to `boolean` values and transforms the flat source paths to the nested target paths [444].

The resulting data is now well-structured, formally specified by a schema, and **AI-ready** [445].

**Integration with Laboratory Robots (XDL Example)** [446]:

For automated execution of experiments using laboratory robots, existing standards like the **Chemical Description Language XDL** [37] are preferred [446, 473].

  - No formal schema is available for XDL 2.0 [474].
  - Relevant sections of the XDL documentation (`⁹https://croningroup.gitlab.io/chemputer/xdl/standard, acc. 25/07/05.`) are copied and pasted into the schema creation prompt of MetaConfigurator to generate an XDL schema [474].
  - Since the XDL structure fundamentally differs from the original data, JSONata cannot provide an adequate mapping [474].
  - The **code generation capabilities** of MetaConfigurator are used to generate Python classes for both the original and the XDL schema [474]. Custom Python code is then written to convert between the structures, resulting in an XDL document instance compatible with the XDL ecosystem, potentially enabling automatic synthesis execution [474, 475].

-----

## V. CONCLUSION

The presented work introduces a **hybrid approach** for schema creation, editing, and instance transformation using LLMs, integrated into the open-source tool **MetaConfigurator** [477].

The method enables intuitive and controlled schema modeling by combining:

  - Modular prompt engineering [478]
  - Targeted context scoping [478]
  - Human-in-the-loop refinement [478]
  - Post-processing safeguards [478]

Mapping rules for transforming large document instances are generated by LLMs but applied in a **deterministic manner** [479]. The chemistry example demonstrates the applicability by translating informal data into machine-readable structures, enabling **FAIR** (Findable, Accessible, Interoperable, and Reusable) [38, 482] data practices and model-driven workflows [480, 482].

**Future Work:** The tool supports MDE tasks such as instance validation, code/documentation generation, mapping discovery, and schema-driven form generation [483]. A future extension is planned for **model-to-model transformations** (e.g., `XSD` to `JSON Schema`) [484].

-----

## ACKNOWLEDGMENT

The authors acknowledge:

  - Esengül Ciftci and Kenichi Endo for inspiring the chemistry application example [486].
  - The assistance of **ChatGPT-4** for editorial suggestions [487].

-----

## REFERENCES

1.  Douglas C. Schmidt, "Guest editor's introduction: Model-driven engineering," Computer, vol. 39, no. 2, pp. 25-31, 2006, doi: `10.1109/MC.2006.58` [489].
2.  G. Pryor, "Managing research data," Facet Publishing, 2012, doi: `10.29085/9781856048910` [490].
3.  S. Wall, "Data management for researchers: Organize, maintain and share your data for research success," Archives and Manuscripts, vol. 45, no. 1, pp. 53-55, 2017, doi: `10.1080/01576895.2017.1279038` [491, 492].
4.  N. Hartl, E. Wößner, and Y. Sure-Vetter, "Nationale Forschungsdateninfrastruktur (NFDI)," Informatik Spektrum, vol. 44, no. 5, pp. 370-373, 2021 [493, 494].
5.  F. Neubauer, P. Bredl, M. Xu, K. Patel, J. Pleiss, and B. Uekermann, "MetaConfigurator: A user-friendly tool for editing structured data files," Datenbank-Spektrum, pp. 1-9, 2024, doi: `10.1007/s13222-024-00472-7` [495].
6.  D. Crockford, "The application/json media type for JavaScript Object Notation (JSON)," Internet Engineering Task Force IETF Request for Comments, 2006 [496].
7.  O. Ben-Kiki, C. Evans, and I. döt Net, "YAML ain't markup language version 1.2," Available: `http://yaml.org/spec/1.2/spec.html`, 2009 [497].
8.  F. Yergeau, T. Bray, J. Paoli, C. M. Sperberg-McQueen, and E. Maler, "Extensible markup language (XML) 1.0," in XML, vol. 1, no. 1, pp. W3C, 2006 [498, 499].
9.  T. Marrs, "JSON at work: Practical data integration for the web," O'Reilly Media, Inc., 2017, pp. 269-286 [500].
10. M.-A. Baazizi, D. Colazzo, G. Ghelli, C. Sartiani, and S. Scherzinger, "An empirical study on the "Usage of Not" in real-world JSON schema documents (long version)," arXiv preprint `arXiv:2107.08677`, 2021, unpublished [501].
11. I. C. Siffa, J. Schäfer, and M. M. Becker, "Adamant: A JSON schema-based metadata editor for research data management workflows," F1000Research, vol. 11, p. 475, Apr. 2022, doi: `10.12688/f1000research.110875.1` [502, 503].
12. F. Neubauer, J. Pleiss, and B. Uekermann, "Data model creation with MetaConfigurator," in Datenbanksysteme für Business, Technologie und Web (BTW 2025), Lecture Notes in Informatics (LNI), vol. P-361, GI, Bonn, 2025, doi: `10.18420/BTW2025-60` [504, 505].
13. M.-A. Baazizi, D. Colazzo, G. Ghelli, and C. Sartiani, "Parametric schema inference for massive JSON datasets," The VLDB Journal, vol. 28. pp. 497-521, 2019, doi: `10.1007/s00778-018-0532-7` [506, 507].
14. W. Spoth, O. Kennedy, Y. Lu, B. Hammerschmidt, and Z. H. Liu, "Reducing ambiguity in JSON schema discovery," in Proc. 2021 Int. Conf. on Management of Data, pp. 1732-1744, 2021 [508, 509].
15. F. Shi, X. Chen, K. Misra, N. Scales, D. Dohan, E. H. Chi, N. Schärli, and D. Zhou, "Large language models can be easily distracted by irrelevant context," in Int. Conf. on Machine Learning, pp. 31210-31227, 2023 [510, 511].
16. R. T. McCoy, S. Yao, D. Friedman, M. Hardy, and T. L. Griffiths, "Embers of autoregression: Understanding large language models through the problem they are trained to solve," arXiv preprint `arXiv:2309.13638`, 2023, unpublished [512].
17. M. Levy, A. Jacoby, and Y. Goldberg, "Same task, more tokens: The impact of input length on the reasoning performance of large language models," arXiv preprint `arXiv:2402,14848`, 2024, unpublished [513].
18. M. J. Mior, "Large language models for JSON schema discovery," arXiv preprint `arXiv:2407.03286`, 2024, unpublished [514].
19. B. Roziere, J. Gehring, F. Gloeckle, S. Sootla, I. Gat, X. E. Tan, Y. Adi, J. Liu, R. Sauvestre, T. Remez, et al., "Code LLAMA: Open foundation models for code," arXiv preprint `arXiv:2308.12950`, 2023, unpublished [515].
20. I. M. Putrama and P. Martinek. "Heterogeneous data integration: Challenges and opportunities." Data in Brief, p. 110853, 2024 [516, 517].
21. B. Ahamed and T. Ramkumar, "Data integration-Challenges, techniques and future directions: A comprehensive study," Indian Journal of Science and Technology, vol. 9, no. 44, pp. 1-9, 2016 [518, 519, 520].
22. F. Z. Rozony, M. N. A. Aktar, M. Ashrafuzzaman, and A. Islam, "A systematic review of big data integration challenges and solutions for heterogeneous data sources," Academic Journal on Business Administration, Innovation & Sustainability, vol. 4, no. 4, pp. 1-18, 2024 [521, 522].
23. R. Hai, C. Quix, and D. Kensche, "Nested schema mappings for integrating JSON," in Conceptual Modeling: Proc. ER 2018, Xi'an, China, Springer, pp. 397-405, 2018 [523].
24. E. Rahm and P. A. Bernstein, "A survey of approaches to automatic schema matching," The VLDB Journal, 10:334-350, 2001. Springer, doi: `10.1007/s007780100057` [524, 525].
25. J. Stanek and D. Killough, "Synthesizing JSON Schema Transformers," arXiv preprint `arXiv:2405.17681`, 2024, unpublished [526].
26. C. Buss, M. Safari, A. Termehchy, S. Lee, and D. Maier, "Towards Scalable Schema Mapping using Large Language Models," arXiv preprint `arXiv:2505.24716`, 2025, unpublished [527].
27. T. Brown, B. Mann, N. Ryder, M. Subbiah, J. D. Kaplan, P. Dhariwal, A. Neelakantan, P. Shyam, G. Sastry, A. Askell, et al., "Language models are few-shot learners," Advances in Neural Information Processing Systems, vol. 33, pp. 1877-1901, 2020, doi: `10.48550/arXiv.2005.14165` [528, 529].
28. J. Achiam. S. Adler, S. Agarwal, L. Ahmad, I. Akkaya, F. L. Aleman, D. Almeida, J. Altenschmidt, S. Altman, S. Anadkat, et al., "GPT-4 technical report," arXiv preprint `arXiv:2303.08774`, 2023, unpublished [529, 530].
29. H. Touvron, T. Lavril, G. Izacard, X. Martinet, M.-A. Lachaux, T. Lacroix, B. Rozière, N. Goyal, E. Hambro, F. Azhar, et al., "LLaMA: Open and efficient foundation language models," arXiv preprint `arXiv:2302.13971`, 2023, unpublished [530, 531].
30. P. Sahoo, A. K. Singh, S. Saha, V. Jain, S. Mondal, and A. Chadha, "A systematic survey of prompt engineering in large language models: Techniques and applications," arXiv preprint `arXiv:2402.07927`, 2024, unpublished [532].
31. A. Kong, S. Zhao, H. Chen, Q. Li, Y. Qin, R. Sun, X. Zhou, E. Wang, and X. Dong. "Better zero-shot reasoning with role-play prompting." arXiv preprint `arXiv:2308.07702`, 2023 [533, 534].
32. S. Mishra, D. Khashabi, C. Baral, Y. Choi, and H. Hajishirzi, "Reframing instructional prompts to GPTk's language," arXiv preprint `arXiv:2109.07830`, 2021, unpublished [535].
33. L. Ouyang, J. Wu, X. Jiang, D. Almeida, C. Wainwright, P. Mishkin, C. Zhang, S. Agarwal, K. Slama, A. Ray, et al., "Training language models to follow instructions with human feedback," Advances in Neural Information Processing Systems, vol. 35, pp. 27730-27744, 2022 [536, 537].
34. Y. Wang, Y. Kordi, S. Mishra, A. Liu, N. A. Smith, D. Khashabi, and H. Hajishirzi, "Self-instruct: Aligning language models with self-generated instructions," arXiv preprint `arXiv:2212.10560`, 2022, unpublished [538].
35. J. Wei, M. Bosma, V. Y. Zhao, K. Guu, A. W. Yu, B. Lester, N. Du, A. M. Dai, and Q. V. Le, "Finetuned language models are zero-shot learners," arXiv preprint `arXiv:2109.01652`, 2021, unpublished [539].
36. M. Seifrid, F. Strieth-Kalthoff, M. Haddadnia, T. C. Wu, E. Alca, L. Bodo, S. Arellano-Rubach, N. Yoshikawa, M. Skreta, and R. Keunen, "Chemspyd: An open-source Python interface for Chemspeed robotic chemistry and materials platforms," Digital Discovery, vol. 3, no. 7. pp. 1319-1326, 2024 [540, 541].
37. S. H. M. Mehr, M. Craven, A. I. Leonov, G. Keenan, and L. Cronin, "A universal system for digitization and automatic execution of the chemical synthesis literature," Science, vol. 370, no. 6512, pp. 101-108, 2020, doi: `10.1126/science.abc2986` [542, 543].
38. M. D. Wilkinson, M. Dumontier, L. J. Aalbersberg, G. Appleton, M. Axton, A. Baak, N. Blomberg, J.-W. Boiten, L. B. da Silva Santos, P. E. Bourne, et al., "The FAIR guiding principles for scientific data management and stewardship," Scientific Data, vol. 3, no. 1, pp. 1-9. 2016 [544, 545, 546].
