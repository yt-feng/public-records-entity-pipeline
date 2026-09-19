# Public Records Entity Pipeline

This repository contains a reproducible pipeline for normalizing public-record entities, explicit relationship claims, and conservative deduplication layers.

The workflow runs on GitHub Actions so the large workbook transformation does not consume a local workstation. It can:

- expand structured relationship edges from the current entity set;
- append only unseen observations;
- rebuild the conservative unique-entity layer;
- materialize a strict scope view; and
- validate IDs, scope boundaries, and workbook integrity.

Run `Process public records` from the Actions tab. Leave relation expansion off for a fast rebuild; turn it on when a fresh public-source expansion is needed. The resulting workbook is uploaded as a workflow artifact and the generated state is committed back to the repository by the workflow.

All records are sourced from public pages or public structured data. A relationship claim or candidate observation is not, by itself, proof of current status or identity. Review the evidence fields before using any record operationally.
