# Incident Assignment Agent Skills

## Objective
The agent reads incident records from an input folder, classifies each incident by severity, assigns an engineer from the service engineering workbook, and writes a structured allocation report to the output folder.

## Workflow
1. Read incidents from the input folder using the first available incident file: incidents.csv
2. Infer severity from the description using a simple keyword-based rubric:
   - High: words such as down, outage, critical, failure, unavailable, emergency.
   - Medium: words such as slow, latency, intermittent, degraded, issue.
   - Low: words such as typo, cosmetic, minor, info, warning.
3. Read the service engineer workbook from the input folder named Service_Engg.xlsx.
4. Rank engineers by:
   - lower current queue count first,
   - higher experience first,
   - specialty match to the incident category when possible.
5. Allocate the most suitable engineer to each incident and write the allocation report to output/allocated_incidents.csv.

## Output Contract
The generated CSV must contain:
- incident_id
- title
- description
- category
- assigned_engineer
- assigned_engineer_experience
- assignment_reason

## Implementation Mapping
This skill document is the human-readable playbook for the agent. The Python implementation in [src/incident_agent.py](src/incident_agent.py) should follow it directly:
- Step 1 maps to the incident loading logic in the script.
- Step 2 maps to the severity classification function.
- Step 3 maps to the engineer workbook reader.
- Step 4 maps to the engineer ranking and assignment logic.
- Step 5 maps to the CSV report writer.

In other words, the skill document defines the workflow, while the script executes it.
