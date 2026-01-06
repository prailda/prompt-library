# Using VBA Excel Prompt Templates - Quick Start Guide

## Introduction

This guide helps you effectively use the VBA Excel prompt templates in this library to generate high-quality code and architecture designs for Excel VBA applications.

## Repository Structure

```
vba-excel/
├── architecture/
│   └── system-prompt.md          # Core architectural guidance
├── code-generation/
│   ├── patterns.md               # Proven code patterns
│   └── hints-directives.md       # Specific coding directives
├── templates/
│   └── application-templates.md  # Ready-to-use application templates
├── examples/
│   └── practical-scenarios.md    # Real-world examples
└── GUIDE.md                       # This file
```

## Quick Start: 3-Step Process

### Step 1: Define Your Requirements

Before using any template, clearly define:

1. **Purpose**: What problem are you solving?
2. **Scope**: What features are essential vs. nice-to-have?
3. **Scale**: How much data? How many users?
4. **Constraints**: Performance requirements, existing systems, skill levels

**Example**:
```
Purpose: Track customer orders and generate invoices
Scope: Essential - CRUD operations, invoice generation; Nice-to-have - Email integration
Scale: ~500 customers, ~100 orders/month
Constraints: Users have basic Excel knowledge, must be maintainable
```

### Step 2: Select Appropriate Resources

#### For Architecture Design:
- Use: `architecture/system-prompt.md`
- When: Starting a new application or redesigning existing one
- Output: High-level architecture, component structure, design decisions

#### For Code Generation:
- Use: `code-generation/patterns.md` + `code-generation/hints-directives.md`
- When: Implementing specific components or features
- Output: Production-ready VBA code following best practices

#### For Quick Start:
- Use: `templates/application-templates.md`
- When: Building common application types (CRUD, reports, etc.)
- Output: Complete starter code that you can customize

#### For Learning:
- Use: `examples/practical-scenarios.md`
- When: Understanding how to solve specific problems
- Output: Example code and patterns for common scenarios

### Step 3: Craft Your Prompt

Combine resources to create an effective prompt:

```
[System Prompt from architecture/system-prompt.md]

Task: Create a customer order management system

Requirements:
- Track customers (name, email, phone, company)
- Record orders (customer, items, quantities, total)
- Generate invoices
- Search and filter orders
- Export to CSV

Additional Context:
- Expected order volume: 100/month
- Users: 3 sales staff with basic Excel knowledge
- Must integrate with existing customer list
- Performance: Must handle 5000+ historical orders

Constraints:
[Key directives from hints-directives.md]
- Use repository pattern for data access
- Implement comprehensive error handling
- Optimize for performance (arrays for bulk operations)
- Use Option Explicit
- Validate all inputs

Please provide:
1. Architecture overview with component diagram
2. Complete implementation of all classes and modules
3. Setup instructions
4. Usage examples
```

## Detailed Usage Patterns

### Pattern 1: Architecture-First Approach

**Use When**: Starting a complex or critical application

**Steps**:
1. Load `architecture/system-prompt.md` as system context
2. Describe your application requirements in detail
3. Ask for architecture design first
4. Review and refine the architecture
5. Request implementation of each component separately
6. Use `patterns.md` and `hints-directives.md` to guide implementation

**Prompt Template**:
```
[Content from architecture/system-prompt.md]

I need to design a VBA application for [PURPOSE].

Business Requirements:
- [Requirement 1]
- [Requirement 2]
- [Requirement 3]

Technical Constraints:
- [Constraint 1]
- [Constraint 2]

Before implementing, please provide:
1. High-level architecture with components and their responsibilities
2. Class diagram showing relationships
3. Data model design
4. Error handling strategy
5. Performance considerations
6. Testing approach

Then, I'll ask you to implement each component following the patterns
in code-generation/patterns.md.
```

### Pattern 2: Template-Based Approach

**Use When**: Building a standard application type quickly

**Steps**:
1. Review `templates/application-templates.md`
2. Select the closest matching template
3. Copy the template code
4. Use specific prompts to customize components

**Prompt Template**:
```
I'm using the Data Entry Application template from application-templates.md.

I need to customize it for [SPECIFIC USE CASE]:

Data Model Changes:
- Add field: [field name] ([type], [validation rules])
- Remove field: [field name]
- Modify field: [changes]

Functionality Changes:
- Add feature: [description]
- Modify behavior: [description]

Please provide the updated code for:
1. clsEntity with new fields
2. Repository methods to handle new fields
3. UserForm layout changes
4. Any additional validation logic

Follow all directives from hints-directives.md especially:
- [Specific directive 1]
- [Specific directive 2]
```

### Pattern 3: Feature Addition

**Use When**: Adding new features to existing code

**Steps**:
1. Review `patterns.md` for the relevant pattern
2. Review `hints-directives.md` for specific guidelines
3. Provide context about existing code
4. Request feature implementation

**Prompt Template**:
```
I have an existing VBA application with [DESCRIPTION].

Current Structure:
- [List key classes and modules]

I need to add: [NEW FEATURE]

Requirements:
- [Specific requirement 1]
- [Specific requirement 2]

Please implement this feature following:
- The [PATTERN NAME] pattern from patterns.md
- These directives from hints-directives.md:
  * [Directive 1]
  * [Directive 2]

Provide:
1. New classes/modules needed
2. Modifications to existing code (minimal changes)
3. Integration points
4. Testing approach
```

### Pattern 4: Code Review and Improvement

**Use When**: Improving existing VBA code

**Prompt Template**:
```
Please review this VBA code against the best practices in
hints-directives.md and patterns.md:

[PASTE YOUR CODE]

Specifically check for:
- Violations of directives D1-D10 in hints-directives.md
- Opportunities to apply patterns from patterns.md
- Performance issues (especially around array usage)
- Error handling gaps
- Security vulnerabilities
- Code organization improvements

For each issue found, provide:
1. Description of the problem
2. Severity (High/Medium/Low)
3. Specific fix with code example
4. Explanation of why the fix is better
```

### Pattern 5: Problem Solving

**Use When**: Solving a specific VBA challenge

**Steps**:
1. Check `examples/practical-scenarios.md` for similar scenarios
2. Identify the relevant scenario or combination
3. Adapt the example to your needs

**Prompt Template**:
```
I need to implement [SPECIFIC FUNCTIONALITY] in VBA.

Context:
- Similar to [SCENARIO NAME] from practical-scenarios.md
- But with these differences: [DIFFERENCES]

Requirements:
- [Requirement 1]
- [Requirement 2]

Constraints:
- [Constraint 1]
- [Constraint 2]

Please provide a solution that:
1. Follows the pattern shown in the similar scenario
2. Adapts it for my specific needs
3. Includes error handling per hints-directives.md
4. Optimizes for performance per hints-directives.md
5. Includes usage examples and test cases
```

## Best Practices for Using This Library

### DO:
✅ Combine multiple resources in your prompts  
✅ Be specific about requirements and constraints  
✅ Reference specific patterns, directives, or examples by name  
✅ Request explanations for why certain approaches are used  
✅ Ask for incremental implementation (one component at a time)  
✅ Request test cases and usage examples  
✅ Specify your VBA skill level for appropriate explanations  

### DON'T:
❌ Use templates without customization  
❌ Skip the architecture phase for complex applications  
❌ Ignore the hints and directives  
❌ Request entire applications without structure  
❌ Forget to specify error handling requirements  
❌ Neglect performance considerations  
❌ Skip validation and testing  

## Common Scenarios and Resource Mapping

| What You Need | Primary Resource | Supporting Resources |
|---------------|-----------------|---------------------|
| Application architecture | architecture/system-prompt.md | patterns.md |
| Class module structure | patterns.md (Pattern 1) | hints-directives.md (D8, D9) |
| Data access layer | patterns.md (Pattern 2) | hints-directives.md (D4, D7) |
| UserForm code | patterns.md (Pattern 3) | hints-directives.md (D10) |
| Error handling | patterns.md (Pattern 4) | hints-directives.md (D7) |
| Configuration management | patterns.md (Pattern 5) | - |
| Complete CRUD app | templates/application-templates.md (Template 1) | All resources |
| Report generation | templates/application-templates.md (Template 2) | patterns.md |
| Specific solutions | examples/practical-scenarios.md | patterns.md, hints-directives.md |

## Example Workflows

### Workflow 1: New Application from Scratch

1. **Architecture Phase**:
   ```
   Use: architecture/system-prompt.md
   Output: Component structure, class diagram, data model
   ```

2. **Implementation Phase** (for each component):
   ```
   Use: patterns.md for structure + hints-directives.md for specifics
   Output: Production-ready code for one component
   ```

3. **Integration Phase**:
   ```
   Use: examples/practical-scenarios.md for integration patterns
   Output: Working application
   ```

4. **Refinement Phase**:
   ```
   Use: hints-directives.md for optimization
   Output: Optimized, polished application
   ```

### Workflow 2: Quick Prototype

1. **Template Selection**:
   ```
   Use: templates/application-templates.md
   Select: Closest matching template
   ```

2. **Customization**:
   ```
   Use: patterns.md for modifications
   Output: Customized template
   ```

3. **Testing**:
   ```
   Use: examples/practical-scenarios.md for test patterns
   Output: Working prototype
   ```

### Workflow 3: Code Improvement

1. **Review**:
   ```
   Use: hints-directives.md for code review criteria
   Output: List of issues
   ```

2. **Refactoring**:
   ```
   Use: patterns.md for better structures
   Output: Refactored code
   ```

3. **Optimization**:
   ```
   Use: hints-directives.md (Performance section)
   Output: Optimized code
   ```

## Customization Tips

### Adapting the System Prompt

The system prompt in `architecture/system-prompt.md` can be customized:

```markdown
Add industry-specific context:
"You specialize in VBA for [INDUSTRY], focusing on [SPECIFIC NEEDS]"

Add compliance requirements:
"All code must comply with [REGULATION] requiring [SPECIFIC PRACTICES]"

Add technology constraints:
"Work within Excel 2016 limitations, no external libraries except [LIST]"
```

### Extending the Patterns

When you discover new patterns in your work:

1. Document the pattern following the structure in `patterns.md`
2. Include: Pattern name, description, code template, key points
3. Add to your personal pattern library
4. Reference your custom patterns in prompts

### Building Custom Templates

Create templates for your specific domain:

1. Start with a template from `application-templates.md`
2. Customize for your industry/use case
3. Add domain-specific validation rules
4. Include your common data structures
5. Save as a reusable template

## Troubleshooting

### Problem: Generated code is too generic

**Solution**: 
- Provide more specific requirements
- Reference specific patterns and directives by name
- Include constraints and edge cases
- Provide example data or scenarios

### Problem: Code doesn't follow best practices

**Solution**:
- Explicitly reference `hints-directives.md` in your prompt
- List specific directives to follow (e.g., "Must follow D1, D4, D7")
- Request code review against the directives

### Problem: Architecture is too complex

**Solution**:
- Specify simplicity as a requirement
- Set constraints on number of classes/modules
- Request incremental approach
- Start with MVP features only

### Problem: Missing error handling

**Solution**:
- Always include in prompt: "Implement error handling per Pattern 4 in patterns.md"
- Reference D7 from hints-directives.md
- Request specific error scenarios to handle

## Advanced Techniques

### Technique 1: Layered Prompting

Build complexity gradually:

```
Prompt 1: Architecture only
Prompt 2: Data layer only
Prompt 3: Business logic layer only
Prompt 4: UI layer only
Prompt 5: Integration and testing
```

Each prompt references previous outputs and relevant library resources.

### Technique 2: Pattern Mixing

Combine multiple patterns:

```
"Implement using:
- Pattern 1 (Class Structure) for entities
- Pattern 2 (Repository) for data access
- Pattern 3 (UserForm Controller) for UI
- Pattern 4 (Error Handling) throughout
- Pattern 5 (Configuration) for settings"
```

### Technique 3: Constraint-Based Generation

Focus on constraints to guide design:

```
"Design must satisfy:
- Performance: Handle 50,000 rows in <5 seconds (Directive D4)
- Maintainability: Maximum 200 lines per procedure (Directive D8)
- Security: Validate all inputs (Directive D10, SEC1-SEC3)
- Testability: All business logic in testable functions"
```

## Measuring Success

Good generated code should:

✅ Compile without errors  
✅ Follow all explicitly requested directives  
✅ Include comprehensive error handling  
✅ Be well-documented (headers and key comments)  
✅ Use appropriate patterns for the task  
✅ Include validation logic  
✅ Be testable (separated concerns)  
✅ Perform efficiently (arrays for bulk ops)  
✅ Match the requested architecture  

## Next Steps

1. **Explore the Resources**: Read through each document to understand available patterns and guidance

2. **Try Simple Examples**: Start with examples from `practical-scenarios.md`

3. **Build a Template**: Create your first application using `application-templates.md`

4. **Refine and Iterate**: Use `hints-directives.md` to improve your code

5. **Create Custom Patterns**: Document patterns specific to your needs

6. **Share and Contribute**: Share successful patterns with your team

## Additional Resources

- **architecture/system-prompt.md**: Comprehensive architectural guidance
- **code-generation/patterns.md**: 5 core patterns with complete code
- **code-generation/hints-directives.md**: 50+ specific directives and hints
- **templates/application-templates.md**: 3 complete application templates
- **examples/practical-scenarios.md**: 6 real-world scenario solutions

## Getting Help

If you're not getting the results you need:

1. Review this guide's troubleshooting section
2. Check that you're using the right resources for your task
3. Make your prompt more specific with concrete examples
4. Break down complex requests into smaller pieces
5. Reference specific patterns, directives, or examples by name

## Conclusion

This library provides a comprehensive foundation for generating high-quality VBA Excel applications. By combining the architectural guidance, proven patterns, specific directives, and practical examples, you can create robust, maintainable, and efficient VBA solutions.

Remember: The key to success is being specific about your requirements and explicitly referencing the patterns and directives you want followed.

Happy coding! 🚀
