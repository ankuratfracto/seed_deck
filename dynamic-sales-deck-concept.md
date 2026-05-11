# Dynamic Client-Specific Sales Deck Concept

## Concept

Create a Fracto workflow that generates a personalized sales deck for every lead before outreach. The workflow takes the lead company name, prospect name, role, segment, and likely document workflows, then produces a client-specific PPTX and message.

The deck should feel prepared for the prospect, not like a generic attachment.

## Required Inputs

| Field | Example | Used For |
| --- | --- | --- |
| `lead_company_name` | `Navata SCS` | Cover slide, footers, CTA, proof points |
| `prospect_name` | `Rahul` | Cover slide, closing slide, outreach message |
| `prospect_role` | `Head of Operations` | Persona-specific value framing |
| `segment` | `road_transport` | Selects relevant deck template |
| `primary_workflow` | `POD review` | Main use case across slides |
| `secondary_workflows` | `trip sheets`, `hire letters`, `invoices` | Supporting use cases |
| `pain_points` | `missing stamps`, `late POD checks` | Problem and impact slides |
| `cta_link` | booking/demo link | Closing slide and message |

## Personalization in the PPT

Recommended dynamic placements:

1. Cover slide
   - `Prepared for {{lead_company_name}}`
   - `Hi {{prospect_name}}, here is how Fracto can help automate {{primary_workflow}} for your transport operations.`

2. Problem slide
   - Replace generic pain points with the most likely issues for that company or segment.
   - Example: `For {{lead_company_name}}, missed stamps, unsigned PODs, and handwritten remarks can slow down billing proof checks.`

3. Capabilities slide
   - Keep Fracto capabilities consistent.
   - Reorder use cases based on `primary_workflow` and `secondary_workflows`.

4. Workflow slide
   - Show the client-specific flow:
     `{{lead_company_name}} uploads PODs -> Fracto checks signatures/stamps/remarks -> exceptions go to ops review -> clean files move to billing.`

5. Segment-fit slide
   - Show similar operators or references from the same segment.
   - Messaging should change by segment: line-haul, 3PL, fleet operator, branch ops, billing proof team.

6. Value slide
   - Use the prospect role to tune outcomes.
   - Ops leader: faster exception routing.
   - Billing leader: cleaner proof readiness.
   - Engineering leader: API-first document automation.

7. Demo scope slide
   - Use a custom demo scope:
     `Demo for {{lead_company_name}}: classify sample PODs, extract fields, detect signature/stamp, flag handwritten remarks.`

8. Closing slide
   - `{{prospect_name}}, want to review this on a few {{lead_company_name}} sample files?`
   - Include `{{cta_link}}`.

## Use Case Selection Logic

Use a rules-based first version:

| Signal | Prioritized Use Cases |
| --- | --- |
| Transport / 3PL / line-haul | POD review, trip sheet checks, billing proof completeness |
| Fleet operator | Driver documents, trip sheets, vehicle movement proof |
| Branch-heavy logistics | Branch upload queues, POD completeness, exception routing |
| Billing or finance persona | Invoice matching, proof completeness, missing signoff detection |
| Operations persona | Queue reduction, exception triage, faster proof readiness |
| Engineering persona | API integration, structured JSON outputs, review workflow hooks |

Later, this can become an LLM step that reads CRM notes, website text, LinkedIn profile, or previous emails and chooses the best deck angle.

## Workflow Design

1. Trigger
   - New lead added, outbound sequence started, or salesperson requests a deck.

2. Enrichment
   - Pull company name, prospect name, role, industry, website, and CRM notes.

3. Personalization
   - Select template: `road_transport`, `quick_commerce`, `mobility`, `education_loans`, etc.
   - Fill variables.
   - Choose top 3-5 use cases.
   - Generate persona-specific message copy.

4. Deck Generation
   - Render personalized HTML deck.
   - Export to PDF.
   - Convert to PPTX or generate PPTX directly.

5. Outreach
   - Attach personalized PPTX.
   - Send message through email, LinkedIn, or WhatsApp workflow.
   - Log generated deck URL and variables back to CRM.

## Recommended Technical Approach

Use HTML as the source of truth for decks and generate final PPTX per prospect.

Why:
- Easier to maintain design.
- Variables are simple to inject.
- Same template can generate PDF and PPTX.
- Avoids fragile PowerPoint macro/add-in dependencies.

Suggested template syntax:

```text
{{lead_company_name}}
{{prospect_name}}
{{prospect_role}}
{{primary_workflow}}
{{pain_points.0}}
{{use_cases.0.title}}
{{cta_link}}
```

## Example Data Payload

```json
{
  "lead_company_name": "Navata SCS",
  "prospect_name": "Rahul",
  "prospect_role": "Head of Operations",
  "segment": "road_transport",
  "primary_workflow": "POD review",
  "secondary_workflows": ["trip sheets", "hire letters", "transport invoices"],
  "pain_points": [
    "manual POD completeness checks",
    "missing signatures and stamps found late",
    "handwritten remarks missed during billing proof review"
  ],
  "use_cases": [
    {
      "title": "POD completeness check",
      "copy": "Detect whether signatures, stamps, and required remarks are present before billing starts."
    },
    {
      "title": "Trip sheet extraction",
      "copy": "Extract trip and movement details into structured fields for ops review."
    },
    {
      "title": "Exception routing",
      "copy": "Send incomplete or uncertain documents to a review queue while clean documents move ahead."
    }
  ],
  "cta_link": "https://fracto.tech/demo"
}
```

## Example Outreach Message

```text
Hi {{prospect_name}},

I put together a short deck for {{lead_company_name}} on how Fracto can automate {{primary_workflow}} and related transport document checks.

The use cases I mapped are around {{pain_points.0}}, {{pain_points.1}}, and {{pain_points.2}}.

Would it make sense to review this on a few sample files from your ops flow?
```

## MVP Scope

Build the first version for the road transport deck:

- Add dynamic cover slide text.
- Add prospect/company-specific problem framing.
- Add dynamic use case cards.
- Add custom demo scope.
- Add personalized closing CTA.
- Generate one PPTX per lead.

The first version can be rules-based. Once the workflow is stable, add LLM-based enrichment to choose better use cases and write sharper prospect-specific copy.
