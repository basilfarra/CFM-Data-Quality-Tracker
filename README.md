# Community Feedback & Complaints Data Quality Tracker

**Excel-based data validation and reporting eligibility for a simulated humanitarian feedback workflow.**

**Author:** Basil Al-Farra  
**Project type:** Independent portfolio project  
**Context:** Monitoring, Evaluation, Accountability and Learning (MEAL) · Information Management  
**Data:** Synthetic practice records only

## Overview

A closed case is not automatically a valid observation for measuring response time. Missing closure dates, inconsistent chronology and unresolved sensitivity classifications can change which records support a performance indicator—and how that indicator should be interpreted.

I developed this project to examine that problem in a simulated community feedback and complaints mechanism (CFM). The workbook preserves the raw dataset, standardises selected fields, removes duplicate case entries, checks data quality and documents eligibility for response-time assessment. It keeps unresolved issues visible and connects the findings to proposed follow-up actions.

The central question is: **Which records can support a defensible reporting result, and what must be disclosed about the records that cannot?**

This is a learning and portfolio exercise. It does not represent deployment by an NGO, management of actual complaints or verified improvements in service delivery.

## Results at a glance

| Measure | Result |
| --- | ---: |
| Raw records | 518 |
| Duplicate entries removed beyond the first occurrence | 18 |
| Unique cases retained | 500 |
| Documented data-quality checks | 13 |
| Cases marked Closed | 289 |
| Closed cases eligible for target assessment | 259 |
| Closed cases excluded from target assessment | 30 |
| Eligible cases within the assumed target | 143 |
| Eligible cases outside the assumed target | 116 |
| Reconciliation difference | 0 |

These figures describe the supplied workbook snapshot. They are not operational performance results. The response-time targets—3 days for Sensitive cases and 15 days for Non-Sensitive cases—are demonstration assumptions, not organisational standards or safeguarding protocols.

## Methodology and decisions

### Preserve the source and standardise consistently

`Raw_Data` retains the original 518 records. `Reference_Lists` contains canonical categories and mappings for variations in Channel, Location, Gender and Status. Cleaning formulas use text normalisation and reference lookups; unmatched values remain visible. Text dates are converted using explicit day, month and year components.

### Retain one record per case

The current workbook retains the first source occurrence of each `Case_ID`. Repeated entries are identical or differ in Channel spelling, spacing or capitalisation. This rule is appropriate to the supplied exercise; it is not a general rule for resolving conflicting case histories.

Stable case identifiers are compared with linked source identifiers through an alignment check. This helps detect a mismatch between a retained case and its source row.

### Keep unresolved defects visible

Standardising a record does not establish that its content is complete or correct. Missing classifications and inconsistent dates remain recorded as exceptions. Proposed actions describe what an authorised team would need to verify at source; they do not imply that field verification occurred.

### Define eligibility before assessing performance

A case enters target assessment only when it is Closed and has a numeric closure date, a numeric non-negative response time and a matched numeric response-time target. Other records receive `NOT ASSESSED` rather than a pass or fail.

Of the 289 closed cases, 30 are excluded using mutually exclusive reasons in the following order:

| Exclusion reason | Cases |
| --- | ---: |
| Missing closure date | 10 |
| Negative response time, with a closure date present | 11 |
| Unresolved sensitivity, with valid dates | 9 |
| Other reason recorded by the workbook | 0 |
| **Total excluded** | **30** |

One case has both a negative response time and missing sensitivity. Adding the overlapping raw counts would produce 31 exclusions. Applying the stated precedence counts that case once and reconciles **259 eligible + 30 excluded = 289 closed**.

## Data-quality findings

The register covers completeness, uniqueness, source alignment, date and status logic, target lookup integrity, calculated-field errors and categorical standardisation.

Selected unresolved findings include:

| Finding | Cases | Reporting implication |
| --- | ---: | --- |
| Missing Sector | 22 | Limits sector-level disaggregation |
| Missing Sensitivity | 15 | Prevents assignment of a response-time target |
| Missing Age_Group | 18 | Limits age-group disaggregation |
| Missing Feedback_Note | 12 | Requires review of whether an appropriate summary is available |
| Closure date earlier than receipt | 11 | Invalidates response-time assessment |
| Closed without a closure date | 10 | Prevents calculation of response time |
| Open or In Progress with a closure date | 8 | Requires verification of status or date |

These counts overlap and must not be added as a count of unique affected cases. The retained dataset still contains unresolved defects; “cleaned” refers to the documented preparation steps, not certification that every record is valid.

## Workbook guide

| Worksheet | Purpose |
| --- | --- |
| `Data_Quality_Check` | Quality register, proposed actions, eligibility reconciliation and management summary |
| `Clean_Data` | Retained cases, prepared fields, response-time calculations and alignment checks |
| `Raw_Data` | Original practice dataset for comparison and traceability |
| `Reference_Lists` | Canonical values, mappings and assumed response-time targets |
| `Build_Spec` | Exercise scope, completed phases and optional future work |

Start with `Data_Quality_Check`, trace the results to `Clean_Data`, then inspect the relevant raw records and reference mappings. Download the Excel workbook and open it in Microsoft Excel to inspect formulas and filters.

The implementation uses formulas including `INDEX`/`MATCH`, `TRIM`, `PROPER`, `DATE`, `IF`, `IFERROR`, `ISNUMBER`, `COUNTIF`, `COUNTBLANK` and `SUMPRODUCT`. Dashboard development, AI-assisted classification, Power Query, VBA and DAX are not part of the completed deliverable.

## Scope and limitations

- All records are synthetic. No actual person, household or organisation is represented.
- Sensitive cases use placeholder text rather than detailed narratives. This is a design boundary of the exercise, not evidence of compliance with an institutional data-protection policy.
- The workbook uses fixed ranges and direct source-row references. Changing the source structure or extending the dataset requires review of formulas, mappings and alignment checks; automatic refresh is not implemented.
- Zero results in the implemented checks apply to their defined scope. They do not establish exhaustive validation, independent assurance or compatibility with every spreadsheet application.
- Suggested programme actions illustrate escalation and verification decisions. They are not records of actions taken by field staff.

## Further development

Possible extensions include a dashboard with explicit reporting denominators, a refresh process based on stable identifiers, and AI-assisted classification of synthetic feedback notes with a manually reviewed comparison sample. These remain future work.

## Learning focus

This project connects practical Excel work with a broader information-systems question: how validation rules, traceable transformations and explicit exclusions shape the reliability of institutional reporting. Its contribution is an inspectable example of those decisions, including the cases where the available data cannot support a performance judgement.
