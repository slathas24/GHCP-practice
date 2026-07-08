# Incident Assignment Agent

This repository contains a GitHub Copilot-oriented incident management agent.

## Responsibilities
- Read incident input files from the input folder.
- Infer severity from descriptions.
- Allocate engineers from the Service_Engg.xlsx workbook.
- Generate the allocation report in the output folder.

## Core guidance
Use the logic defined in [SKILLS.md](SKILLS.md) and the implementation in [src/incident_agent.py](src/incident_agent.py).
Store the generated output in output folder with the date and timestamp appended to the name of the report file 

