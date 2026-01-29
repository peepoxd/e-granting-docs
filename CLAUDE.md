# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Model Preference

## Project Overview

PM Exam Assistant for ThaiHealth (สสส.) Project Manager IT position exam preparation. Creates one-page A4 documents with diagrams for two exam types:
- **ข้อ 1**: System architecture design for สสส. project management

## Commands

```bash
# Generate Word documents
python3 create_word_docs.py

# Requires python-docx
pip install python-docx
```

1

Required MCP servers for diagram generation:

```json
{
  "mcpServers": {
    "mermaid": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-mermaid"]
    },
    "antv-chart": {
      "command": "npx",
      "args": ["-y", "@antv/mcp-server-chart"]
    }
  }
}
```

Diagram mapping:
- Architecture Diagrams → Mermaid
- Process Flows → AntV `generate_flow_diagram`
- Mind Maps → AntV `generate_mind_map`

## Document Specifications

**Output format:**
- A4 Portrait, exactly 1 page
- Font: TH Sarabun New 14pt (12-13pt if content overflows)
- Margin: 1 inch (2.54 cm) or 0.8 inch for tighter fit
- Diagram:text ratio ≈ 50:50

**Language:** Formal Thai (ภาษาราชการ) - avoid slang or informal expressions

## Document Structure

### ข้อ 1 - System Design (สสส.)
```
[50% - Diagram] Title + System Architecture Diagram
[50% - Content] Components (15%) + Features (15%) + Tech Stack (10%) + Benefits (10%)
```

Key สสส. processes: พัฒนาข้อเสนอ → อนุมัติ → เบิกเงิน → ส่งงาน → ปิดโครงการ

## Reference Systems

Thai government systems for context: e-GP (procurement), GFMIS (finance), PMQA (evaluation), e-Budget
