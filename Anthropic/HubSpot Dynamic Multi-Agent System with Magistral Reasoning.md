# HubSpot Dynamic Multi-Agent System with Magistral Reasoning

[![Open In Colab](https://camo.githubusercontent.com/96889048f8a9014fdeba2a891f97150c6aac6e723f5190236b10215a97ed41f3/68747470733a2f2f636f6c61622e72657365617263682e676f6f676c652e636f6d2f6173736574732f636f6c61622d62616467652e737667)](https://colab.research.google.com/github/mistralai/cookbook/blob/main/mistral/agents/non_framework/hubspot_dynamic_multi_agent/hubspot_dynamic_multi_agent_system.ipynb)

This cookbook demonstrates the power of Magistral reasoning model combined with HubSpot CRM integration to create an intelligent, multi-agent system that can understand complex business queries and execute sophisticated CRM operations automatically.  
The system transforms natural language business questions into actionable insights and automated CRM updates, showcasing how advanced AI reasoning can streamline sales operations and strategic decision-making.

![Sample Demo](https://raw.githubusercontent.com/mistralai/cookbook/33afd34a9b8a724e93f5e8501b220189de896abe/mistral/agents/non_framework/hubspot_dynamic_multi_agent/assets/demo.gif)

## Problem Statement

### Traditional CRM Challenges

Modern sales and marketing teams face several critical challenges when working with CRM systems like HubSpot:

**Manual Data Analysis:** Teams spend hours manually analyzing deals, contacts, and companies to extract insights **Complex Query Processing:** Business stakeholders struggle to get answers to multi-faceted questions that require data from multiple CRM objects **Strategic Planning:** Market analysis and expansion planning requires combining CRM data with business intelligence in ways that aren't natively supported

### Sample Query

"Assign priorities to all deals based on deal value"

These queries require:

- Understanding business context
- Analyzing multiple data sources
- Applying business logic
- Generating actionable recommendations
- Sometimes updating CRM records automatically

## Solution Architecture

**Core Innovation:** Magistral Reasoning + HubSpot Integration + Multi-Agent Orchestration

Our solution combines **Mistral's Magistral reasoning model** with **HubSpot's comprehensive CRM API** through a sophisticated multi-agent system that can:

- **Understand** complex business queries using Magistral's advanced reasoning capabilities
- **Plan** multi-step execution strategies with dynamically created specialized agents
- **Execute** both data analysis and CRM updates through coordinated agent workflows
- **Synthesize** results into actionable business insights with strategic recommendations

### AgentOrchestrator

**Master coordinator** that manages the entire multi-agent workflow and HubSpot integration. Orchestrates the complete flow from query analysis through sub-agent execution to final synthesis, while managing agent lifecycle and data connectivity.

### LeadAgent

Powered by **Magistral reasoning model** with `<think>` pattern processing, the Lead Agent performs sophisticated query analysis to understand business intent, determine data requirements, and create detailed execution plans specifying which sub-agents to create dynamically.

### Dynamic Sub-Agents

Sub-agents are **created on-the-fly** based on specific query requirements - not pre-defined templates. Each agent is dynamically generated with specialized roles (e.g., priority_calculator, market_analyzer, deals_updater), specific tasks, and targeted data access patterns using **Mistral Small** for fast execution.

### HubSpot API Connector

Dedicated connector providing comprehensive access to CRM data and operations:

- **Property Discovery:** Automatically maps all available HubSpot fields and valid values
- **Data Fetching:** Retrieves deals, contacts, and companies with full property sets
- **Batch Updates:** Efficiently updates multiple records in batches of 100

### SynthesisAgent

**Final orchestrator** that combines all sub-agent results into coherent, actionable business insights using **Mistral Small**. Transforms technical agent outputs into user-friendly responses with strategic recommendations and next steps.

![Solution Architecture](https://raw.githubusercontent.com/mistralai/cookbook/33afd34a9b8a724e93f5e8501b220189de896abe/mistral/agents/non_framework/hubspot_dynamic_multi_agent/assets/solution_architecture.png)

#### Installation

We need `hubspot-api-client` and `mistralai` packages for the demonstration.

!pip install hubspot-api-client=="12.0.0" mistralai=="1.9.3"

Collecting hubspot-api-client
  Downloading hubspot_api_client-12.0.0-py3-none-any.whl.metadata (9.3 kB)
Collecting mistralai
  Downloading mistralai-1.9.3-py3-none-any.whl.metadata (37 kB)
Requirement already satisfied: requests>=2.31.0 in /usr/local/lib/python3.11/dist-packages (from hubspot-api-client) (2.32.3)
Requirement already satisfied: urllib3<3.0,>=1.15 in /usr/local/lib/python3.11/dist-packages (from hubspot-api-client) (2.5.0)
Requirement already satisfied: six<2.0,>=1.10 in /usr/local/lib/python3.11/dist-packages (from hubspot-api-client) (1.17.0)
Requirement already satisfied: certifi>=2023.1.1 in /usr/local/lib/python3.11/dist-packages (from hubspot-api-client) (2025.7.14)
Requirement already satisfied: python-dateutil>=2.8.2 in /usr/local/lib/python3.11/dist-packages (from hubspot-api-client) (2.9.0.post0)
Collecting eval-type-backport>=0.2.0 (from mistralai)
  Downloading eval_type_backport-0.2.2-py3-none-any.whl.metadata (2.2 kB)
Requirement already satisfied: httpx>=0.28.1 in /usr/local/lib/python3.11/dist-packages (from mistralai) (0.28.1)
Requirement already satisfied: pydantic>=2.10.3 in /usr/local/lib/python3.11/dist-packages (from mistralai) (2.11.7)
Requirement already satisfied: typing-inspection>=0.4.0 in /usr/local/lib/python3.11/dist-packages (from mistralai) (0.4.1)
Requirement already satisfied: anyio in /usr/local/lib/python3.11/dist-packages (from httpx>=0.28.1->mistralai) (4.9.0)
Requirement already satisfied: httpcore==1.* in /usr/local/lib/python3.11/dist-packages (from httpx>=0.28.1->mistralai) (1.0.9)
Requirement already satisfied: idna in /usr/local/lib/python3.11/dist-packages (from httpx>=0.28.1->mistralai) (3.10)
Requirement already satisfied: h11>=0.16 in /usr/local/lib/python3.11/dist-packages (from httpcore==1.*->httpx>=0.28.1->mistralai) (0.16.0)
Requirement already satisfied: annotated-types>=0.6.0 in /usr/local/lib/python3.11/dist-packages (from pydantic>=2.10.3->mistralai) (0.7.0)
Requirement already satisfied: pydantic-core==2.33.2 in /usr/local/lib/python3.11/dist-packages (from pydantic>=2.10.3->mistralai) (2.33.2)
Requirement already satisfied: typing-extensions>=4.12.2 in /usr/local/lib/python3.11/dist-packages (from pydantic>=2.10.3->mistralai) (4.14.1)
Requirement already satisfied: charset-normalizer<4,>=2 in /usr/local/lib/python3.11/dist-packages (from requests>=2.31.0->hubspot-api-client) (3.4.2)
Requirement already satisfied: sniffio>=1.1 in /usr/local/lib/python3.11/dist-packages (from anyio->httpx>=0.28.1->mistralai) (1.3.1)
Downloading hubspot_api_client-12.0.0-py3-none-any.whl (4.3 MB)
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 4.3/4.3 MB 43.8 MB/s eta 0:00:00
Downloading mistralai-1.9.3-py3-none-any.whl (426 kB)
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 426.3/426.3 kB 17.6 MB/s eta 0:00:00
Downloading eval_type_backport-0.2.2-py3-none-any.whl (5.8 kB)
Installing collected packages: eval-type-backport, hubspot-api-client, mistralai
Successfully installed eval-type-backport-0.2.2 hubspot-api-client-12.0.0 mistralai-1.9.3

#### Imports

import requests import json from mistralai import Mistral, ThinkChunk, TextChunk from datetime import datetime, timedelta from typing import Dict, List, Any, Optional import re

#### Setup API Keys

HUBSPOT_API_KEY = "" # Replace with your HubSpot API key MISTRAL_API_KEY = "" # Get it from https://console.mistral.ai/api-keys

#### Setup MistralAI Client

mistral_client = Mistral(api_key=MISTRAL_API_KEY)

#### HubSpot API connector

- `get_data`: Fetches CRM data from HubSpot API, Retrieves deals, contacts, and companies data for the analysis.
    
- `batch_update`: Performs batch updates to HubSpot records, the updates and writes them back to HubSpot in efficient batches of 100 records.
    
- `get_properties`: Automatically fetches and formats all HubSpot deal, contact, and company properties, including valid values and dropdown options, so agents can update data reliably without errors.
    

class HubSpotConnector: """Handles all HubSpot API operations""" def __init__(self, api_key: str): self.api_key = api_key self.base_url = "https://api.hubapi.com/crm/v3/objects" self.headers = {"Authorization": f"Bearer {api_key}"} def get_properties(self) -> Dict: """Load all HubSpot properties""" print("📡 HubSpotConnector: Loading properties...") properties = {} for obj_type in ['deals', 'contacts', 'companies']: url = f"https://api.hubapi.com/crm/v3/properties/{obj_type}" response = requests.get(url, headers=self.headers) if response.status_code == 200: props = response.json()['results'] prop_list = [] for prop in sorted(props, key=lambda x: x['name']): prop_str = f"'{prop['name']}' - {prop['label']}" if 'options' in prop and prop['options']: valid_values = [opt['value'] for opt in prop['options']] prop_str += f" | Valid values: {valid_values}" prop_list.append(prop_str) properties[obj_type] = prop_list print(f"✅ HubSpotConnector: Loaded properties for {len(properties)} object types") return properties def get_data(self, object_type: str) -> List[Dict]: """Fetch data from HubSpot""" print(f"📡 HubSpotConnector: Fetching {object_type} data...") url = f"{self.base_url}/{object_type}" params = {"limit": 100} all_data = [] while url: response = requests.get(url, headers=self.headers, params=params) if response.status_code != 200: raise Exception(f"HubSpot API error: {response.text}") data = response.json() all_data.extend(data.get("results", [])) url = data.get("paging", {}).get("next", {}).get("link") params = {} print(f"✅ HubSpotConnector: Loaded {len(all_data)} {object_type}") return all_data def batch_update(self, updates: Dict) -> None: """Perform batch updates to HubSpot""" for object_type, update_list in updates.items(): if not update_list: continue print(f"📡 HubSpotConnector: Updating {len(update_list)} {object_type}...") url = f"{self.base_url}/{object_type}/batch/update" headers = {**self.headers, "Content-Type": "application/json"} # Process in batches of 100 for i in range(0, len(update_list), 100): batch = update_list[i:i+100] payload = {"inputs": batch} response = requests.post(url, headers=headers, json=payload) if response.status_code not in [200, 202]: raise Exception(f"HubSpot update error: {response.text}") print(f"✅ HubSpotConnector: {object_type} updates completed")

#### Magistral (reasoning) and Mistral small LLM functions

- `magistral_reasoning`: Uses Magistral reasoning model for complex query analysis and execution planning with thinking process.
    
- `mistral_small_execution`: Uses Mistral Small model for sub-agent task execution.
    

def magistral_reasoning(prompt: str) -> Dict[str, str]: """Use reasoning model for query analysis and planning""" response = mistral_client.chat.complete( model="magistral-medium-latest", messages=[{"role": "user", "content": prompt}] ) content = response.choices[0].message.content reasoning = "" conclusion = "" for r in content: if isinstance(r, ThinkChunk): reasoning = r.thinking[0].text elif isinstance(r, TextChunk): conclusion = r.text return { "reasoning": reasoning, "conclusion": conclusion } def mistral_small_execution(prompt: str) -> str: """Use Mistral Small for content generation""" response = mistral_client.chat.complete( model="mistral-small-latest", messages=[{"role": "user", "content": prompt}] ) return response.choices[0].message.content

#### LeadAgent

**Powered by Magistral reasoning model** for sophisticated query analysis and execution planning

- `analyze_query`: Uses Magistral's `<think>` pattern to understand business intent, determine data requirements, and create detailed execution plans with dynamic sub-agent specifications
- Determines whether queries require read-only analysis or write-back operations to HubSpot

class LeadAgent: """Lead Agent powered by Magistral reasoning model for query analysis and planning""" def __init__(self, hubspot_properties): self.hubspot_properties = hubspot_properties self.name = "LeadAgent" def analyze_query(self, query: str) -> Dict: """Analyze query using Magistral reasoning and create execution plan""" print(f"🧠 {self.name}: Analyzing query with Magistral reasoning...") analysis_prompt = f""" Analyze this HubSpot query and create a detailed execution plan based on different hubspot properties provided by following the shared rules: HUBSPOT_PROPERTIES: {self.hubspot_properties} QUERY: {query} RULES: 1. What is the user asking for? 2. Is this a read-only query or does it require HubSpot updates? 3. What sub-agents are needed to accomplish this? 4. What HubSpot data is required? 5. What's the execution sequence? 6. What should be the final output format? 7. Query can also be combination of read-only and write-back. 8. Query is read-only if it requires data read from HubSpot. 9. Query is write-back if it requires an update to existing values or writing/ assigning new values. 10. In the final conclusion just give only one JSON string nothing else. I don't need any explanation. Provide a JSON execution plan with: {{ "sub_agents": [ {{ "name": "agent_name", "task": "specific task description", "task_type": "read_only" or "write_back", "input_data": ["deals", "contacts", "companies"], "output_format": "expected output" }} ] }} """ # Use existing magistral_reasoning function analysis = magistral_reasoning(analysis_prompt) try: # Extract JSON execution plan from conclusion json_match = re.search(r'\{.*\}', analysis["conclusion"], re.DOTALL) if json_match: execution_plan = json.loads(json_match.group(0)) else: raise ValueError("No JSON found in analysis") except Exception as e: print(f"⚠️ {self.name}: JSON parsing failed, using fallback plan") execution_plan = { "sub_agents": [{ "name": "general_analyzer", "task": query, "task_type": "read_only", "input_data": ["deals", "contacts", "companies"], "output_format": "summary" }] } print(f"✅ {self.name}: Plan created - {len(execution_plan['sub_agents'])} sub-agents needed") return { "reasoning": analysis["reasoning"], "execution_plan": execution_plan, "conclusion": analysis["conclusion"] }

#### SubAgent

**Dynamic agents created on-the-fly** based on query complexity and requirements

- `execute`: Uses Mistral Small for fast task execution including data analysis, business logic application, and CRM updates
- Specialized roles generated automatically (e.g., priority_calculator, market_analyzer, deals_updater)
- Handles both read-only operations and write-back operations with proper HubSpot property validation

class SubAgent: """Dynamic Sub-Agent created on-the-fly for specific tasks""" def __init__(self, name: str, task: str, task_type: str, input_data: List[str], output_format: str): self.name = name self.task = task self.task_type = task_type self.input_data = input_data self.output_format = output_format def execute(self, data: Dict, properties_context: str, hubspot_updater=None) -> Dict: """Execute the assigned task""" print(f"🤖 {self.name} ({self.task_type}): Executing task...") if self.task_type == 'read_only': agent_prompt = f""" You are a {self.name} agent. TASK: {self.task} AVAILABLE HUBSPOT PROPERTIES: {properties_context} DATA AVAILABLE: {json.dumps(data, indent=2)} OUTPUT FORMAT: {self.output_format} Provide your analysis only based on the available data. """ else: # write_back agent_prompt = f""" You are a {self.name} agent. TASK: {self.task} AVAILABLE HUBSPOT PROPERTIES: {properties_context} DATA AVAILABLE: {json.dumps(data, indent=2)} CRITICAL: Use exact HubSpot property names from the list above in your JSON output. OUTPUT FORMAT: JSON format with the properties to be written to HubSpot Provide updates using exact HubSpot property names. """ # Use existing mistral_small_execution function result = mistral_small_execution(agent_prompt) # Handle write-back operations if self.task_type == 'write_back' and hubspot_updater: try: json_match = re.search(r'\{.*\}', result, re.DOTALL) if json_match: updates = json.loads(json_match.group(0)) hubspot_updater.batch_update(updates) print(f"✅ {self.name}: Successfully updated HubSpot records") except Exception as e: print(f"❌ {self.name}: Update failed - {str(e)}") return {"status": "error", "error": str(e), "raw_result": result} print(f"✅ {self.name}: Task completed successfully") return {"status": "success", "result": result}

#### SynthesisAgent

**Final orchestrator** that combines all sub-agent results into coherent business insights

- `synthesize`: Uses Mistral Small to create user-friendly responses with actionable recommendations and next steps
- Transforms technical agent outputs into executive-ready summaries and strategic guidance

class SynthesisAgent: """Final agent to synthesize all results into user-friendly response""" def __init__(self): self.name = "SynthesisAgent" def synthesize(self, query: str, sub_agent_results: List[Dict], execution_plan: Dict) -> str: """Combine all sub-agent results into final answer""" print(f"🔄 {self.name}: Synthesizing results from {len(sub_agent_results)} agents...") # Prepare context from all sub-agent results results_context = "" for result in sub_agent_results: results_context += f"\n{result['agent'].upper()} ({result['task_type']}):\n" if result['result']['status'] == 'success': results_context += f"{result['result']['result']}\n" else: results_context += f"Error: {result['result'].get('error', 'Unknown error')}\n" results_context += "---\n" synthesis_prompt = f""" You are a final synthesizer agent. Create a comprehensive, user-friendly response based on all sub-agent results. ORIGINAL QUERY: {query} SUB-AGENT RESULTS: {results_context} TASK: Synthesize all the above results into a clear, actionable response for the user. Guidelines: 1. Start with a direct answer to the user's query 2. Include key insights and findings 3. If updates were made, summarize what was changed 4. Provide actionable next steps if relevant 5. Keep it concise but comprehensive 6. Use a professional but friendly tone Provide the final synthesized response: """ # Use existing mistral_small_execution function final_answer = mistral_small_execution(synthesis_prompt) print(f"✅ {self.name}: Final answer synthesized") return final_answer

#### AgentOrchestrator

**Master coordinator** that manages the entire multi-agent workflow and HubSpot integration

- `process_query`: Orchestrates the complete flow from query analysis through sub-agent execution to final synthesis
- Manages agent lifecycle, data flow between agents, and HubSpot connectivity
- Provides rich logging and monitoring of the multi-agent process

class AgentOrchestrator: """Main orchestrator that coordinates all agents""" def __init__(self, hubspot_api_key: str, mistral_api_key: str): # Initialize global mistral client for existing functions global mistral_client mistral_client = Mistral(api_key=mistral_api_key) # Initialize HubSpot connector self.hubspot_connector = HubSpotConnector(hubspot_api_key) # Load HubSpot data and properties self.hubspot_properties = self.hubspot_connector.get_properties() self.hubspot_data = { "deals": self.hubspot_connector.get_data("deals"), "contacts": self.hubspot_connector.get_data("contacts"), "companies": self.hubspot_connector.get_data("companies") } # Initialize agents self.lead_agent = LeadAgent(self.hubspot_properties) self.synthesis_agent = SynthesisAgent() self.active_sub_agents = [] print(f"🚀 AgentOrchestrator: System initialized with {sum(len(data) for data in self.hubspot_data.values())} HubSpot records") def process_query(self, query: str) -> Dict: """Main method to process user queries through multi-agent workflow""" print(f"\n🎯 Processing Query: {query[:100]}...") print("=" * 80) # Step 1: Lead Agent analyzes query using Magistral reasoning analysis = self.lead_agent.analyze_query(query) execution_plan = analysis["execution_plan"] # Step 2: Create and execute sub-agents dynamically sub_agent_results = [] self.active_sub_agents = [] for agent_config in execution_plan["sub_agents"]: # Create sub-agent dynamically sub_agent = SubAgent( name=agent_config["name"], task=agent_config["task"], task_type=agent_config["task_type"], input_data=agent_config["input_data"], output_format=agent_config["output_format"] ) self.active_sub_agents.append(sub_agent) # Prepare data and context for this sub-agent agent_data = {data_type: self.hubspot_data.get(data_type, []) for data_type in agent_config["input_data"]} # Build properties context properties_context = "" for data_type in agent_config["input_data"]: if data_type in self.hubspot_properties: properties_context += f"\n{data_type.upper()} PROPERTIES:\n" properties_context += "\n".join(self.hubspot_properties[data_type]) properties_context += "\n" # Execute sub-agent using mistral_small_execution result = sub_agent.execute(agent_data, properties_context, self.hubspot_connector) sub_agent_results.append({ "agent": sub_agent.name, "task_type": sub_agent.task_type, "result": result }) # Step 3: Synthesis Agent creates final answer using mistral_small_execution final_answer = self.synthesis_agent.synthesize(query, sub_agent_results, execution_plan) print("=" * 80) print("✨ Query processing completed!") return { "query": query, "reasoning": analysis["reasoning"], "execution_plan": execution_plan, "sub_agent_results": sub_agent_results, "active_agents": [agent.name for agent in self.active_sub_agents], "final_answer": final_answer }

#### Initialize the multi-agent system

orchestrator = AgentOrchestrator( hubspot_api_key=HUBSPOT_API_KEY, mistral_api_key=MISTRAL_API_KEY )

📡 HubSpotConnector: Loading properties...
✅ HubSpotConnector: Loaded properties for 3 object types
📡 HubSpotConnector: Fetching deals data...
✅ HubSpotConnector: Loaded 5 deals
📡 HubSpotConnector: Fetching contacts data...
✅ HubSpotConnector: Loaded 5 contacts
📡 HubSpotConnector: Fetching companies data...
✅ HubSpotConnector: Loaded 5 companies
🚀 AgentOrchestrator: System initialized with 15 HubSpot records

#### Test Queries

##### Query-1

query = "Assign priorities to all deals based on deal value." result = orchestrator.process_query(query)

🎯 Processing Query: Assign priorities to all deals based on deal value....
================================================================================
🧠 LeadAgent: Analyzing query with Magistral reasoning...
✅ LeadAgent: Plan created - 2 sub-agents needed
🤖 deal_data_fetcher (read_only): Executing task...
✅ deal_data_fetcher: Task completed successfully
🤖 priority_assigner (write_back): Executing task...
📡 HubSpotConnector: Updating 5 deals...
✅ HubSpotConnector: deals updates completed
✅ priority_assigner: Successfully updated HubSpot records
✅ priority_assigner: Task completed successfully
🔄 SynthesisAgent: Synthesizing results from 2 agents...
✅ SynthesisAgent: Final answer synthesized
================================================================================
✨ Query processing completed!

##### HubSpot status before updation

![HubSpot Status Before Updation](https://raw.githubusercontent.com/mistralai/cookbook/33afd34a9b8a724e93f5e8501b220189de896abe/mistral/agents/non_framework/hubspot_dynamic_multi_agent/assets/hubspot_status_before_updation.png)

##### HubSpot status after updation

![HubSpot Status After Updation](https://raw.githubusercontent.com/mistralai/cookbook/33afd34a9b8a724e93f5e8501b220189de896abe/mistral/agents/non_framework/hubspot_dynamic_multi_agent/assets/hubspot_status_after_updation.png)

##### Dynamically Created Agents

agents = '\n'.join([f"{i + 1}. {agent}" for i, agent in enumerate(result['active_agents'])]) display(Markdown(agents))

1. deal_data_fetcher
2. priority_assigner

##### Answer

from IPython.display import display, Markdown, Latex display(Markdown(result['final_answer']))

**Final Synthesized Response:**

**Deal Priorities Assigned Based on Value:** We’ve successfully assigned priorities to all deals based on their values. Here’s the breakdown:

1. **High Priority (1 deal):**
    
    - **Enterprise Ltd - Platform Integration** ($200,000)
    - _Reason:_ Value exceeds $200,000.
2. **Medium Priority (2 deals):**
    
    - **Global Solutions - Sales Training** ($125,000)
    - **Acme Corp - CRM Implementation** ($85,000)
    - _Reason:_ Values fall between 199,999.
3. **Low Priority (2 deals):**
    
    - **StartupXYZ - Consulting Services** ($25,000)
    - **TechStart Inc - Marketing Automation** ($45,000)
    - _Reason:_ Values below $50,000.

**Key Insights:**

- The highest-value deal ($200,000) is prioritized as _high_.
- Medium-priority deals are in the mid-range value tier.
- Low-priority deals are smaller in value but may still be important for long-term growth.

**Next Steps:**

- Review the priorities to ensure alignment with business goals.
- Focus resources on high/medium-priority deals to maximize revenue impact.
- Monitor low-priority deals for potential upsell opportunities.

Let us know if you’d like to adjust the criteria or need further analysis!

---

_All updates have been applied to your CRM system._

##### Query-2

query = """We're considering expanding into three new industry verticals and need comprehensive market intelligence to inform our go-to-market strategy. Analyze our current customer base to identify patterns in successful account profiles, understand the characteristics that predict customer success, and use these insights to evaluate market opportunities. The analysis should identify which industries show the strongest fit with our solution, what use cases resonate most effectively, and what competitive landscape we would face. Develop ideal customer profiles for each target market, estimate market size and penetration potential, and create a prioritized market entry strategy with resource requirements and timeline projections for successful market penetration.""" result = orchestrator.process_query(query)

🎯 Processing Query: We're considering expanding into three new industry verticals and need comprehensive market
    inte...
================================================================================
🧠 LeadAgent: Analyzing query with Magistral reasoning...
✅ LeadAgent: Plan created - 6 sub-agents needed
🤖 Customer Success Analysis Agent (read_only): Executing task...
✅ Customer Success Analysis Agent: Task completed successfully
🤖 Market Opportunity Analysis Agent (read_only): Executing task...
✅ Market Opportunity Analysis Agent: Task completed successfully
🤖 Competitive Landscape Analysis Agent (read_only): Executing task...
✅ Competitive Landscape Analysis Agent: Task completed successfully
🤖 Ideal Customer Profile Agent (read_only): Executing task...
✅ Ideal Customer Profile Agent: Task completed successfully
🤖 Market Size and Penetration Agent (read_only): Executing task...
✅ Market Size and Penetration Agent: Task completed successfully
🤖 Market Entry Strategy Agent (read_only): Executing task...
✅ Market Entry Strategy Agent: Task completed successfully
🔄 SynthesisAgent: Synthesizing results from 6 agents...
✅ SynthesisAgent: Final answer synthesized
================================================================================
✨ Query processing completed!

##### Dynamically Created Agents

agents = '\n'.join([f"{i + 1}. {agent}" for i, agent in enumerate(result['active_agents'])]) display(Markdown(agents))

1. Customer Success Analysis Agent
2. Market Opportunity Analysis Agent
3. Competitive Landscape Analysis Agent
4. Ideal Customer Profile Agent
5. Market Size and Penetration Agent
6. Market Entry Strategy Agent

##### Answer

from IPython.display import display, Markdown, Latex display(Markdown(result['final_answer']))

**Comprehensive Market Expansion Strategy for Three New Industry Verticals**

**Key Findings & Recommendations:**

1. **Ideal Customer Profiles & Market Fit**
    
    - **Top Industries for Expansion:**
        - **Technology/Software (SaaS, Fintech, Enterprise Solutions)** – Highest alignment with your current successful accounts (e.g., TechStart Inc, Global Solutions).
        - **Enterprise Solutions** – Strong fit for large-scale deals (e.g., Enterprise Ltd).
        - **Corporate/Industrial Services** – Potential in B2B industrial or supply chain solutions (e.g., Acme Corp).
    - **Ideal Customer Profiles:**
        - **Enterprise Market:** Large organizations needing scalable solutions (e.g., Robert Davis at Enterprise Ltd).
        - **Tech Startups:** Fast-growing companies requiring agile tools (e.g., Sarah Wilson at TechStart Inc).
        - **Global Services:** Consulting/outsourcing firms with cross-border needs (e.g., Michael Johnson at Global Solutions).
2. **Competitive Landscape**
    
    - Limited data, but inferred competitors include:
        - **TechStart Inc & StartupXYZ** (direct competitors in SaaS/startup space).
        - **Global Solutions & Enterprise Ltd** (broader enterprise/business solutions).
    - **Recommendation:** Enrich data with industry classifications, revenue, and engagement metrics to refine competitive positioning.
3. **Market Size & Penetration Potential**
    
    - Current dataset is too limited (only 5 companies) for accurate sizing.
    - **Actionable Step:** Expand data collection to include:
        - Industry, company size, geographic location.
        - Engagement metrics (emails, meetings, website activity).
4. **Prioritized Market Entry Strategy**
    
    - **High-Priority Accounts (Immediate Focus):**
        - **Global Solutions ($125K, Lead Captured)** & **Acme Corp ($85K, Lead Nurtured)** – Assign dedicated AEs, customize proposals, and accelerate deal closure.
    - **Medium-Priority (TechStart Inc, $45K)** – Standardized demos + email nurturing.
    - **Low-Priority (StartupXYZ, $25K)** – Automated nurturing until engagement increases.
    - **Timeline:**
        - **Discovery/Qualification:** 1–2 weeks.
        - **Proposal/Demo:** 1–2 weeks.
        - **Negotiation/Close:** 1–3 weeks.
        - **Onboarding:** 2–4 weeks.
5. **Resource Allocation**
    
    - **Sales:** Dedicate AEs to high-value accounts; automate nurturing for low-priority leads.
    - **Marketing:** Targeted campaigns for medium-priority accounts (e.g., case studies, retargeting ads).

**Next Steps:**

1. **Enrich Data:** Add industry, revenue, and engagement metrics to refine targeting.
2. **Launch Pilot Campaigns:** Test strategies with high-priority accounts (Global Solutions, Acme Corp).
3. **Monitor Metrics:** Track conversion rates, time-to-close, and customer satisfaction.

**Conclusion:** Focus on **technology/enterprise solutions** first, prioritize high-value accounts, and scale based on engagement data. With refined data and targeted execution, you can efficiently penetrate these markets.

Would you like to dive deeper into any specific area (e.g., competitive analysis, customer success tactics)?