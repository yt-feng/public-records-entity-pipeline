# Public Records Entity Pipeline

This repository contains a reproducible pipeline for normalizing public-record entities, explicit relationship claims, and conservative deduplication layers.

The workflow runs on GitHub Actions so the large workbook transformation does not consume a local workstation. Final XLSX materialization uses a constant-memory streaming writer on the runner, so the public workbook is rebuilt without requiring a local desktop process. It can:

- expand structured relationship edges from the current entity set;
- append only unseen observations;
- rebuild the conservative unique-entity layer;
- materialize a strict scope view; and
- validate IDs, scope boundaries, and workbook integrity.

Run `Process public records` from the Actions tab. Automatic relation waves use fast mode: the append-only compressed JSON state and compact entity index are updated on each batch, while the large XLSX is materialized when the unique target is reached. Set `fast_mode=false` when an immediate workbook rebuild is required, or set `materialize_workbook=true` for an explicit materialization run. The resulting workbook is uploaded as a workflow artifact when materialized, and generated state is committed back to the repository by the workflow.

All records are sourced from public pages or public structured data. A relationship claim or candidate observation is not, by itself, proof of current status or identity. Review the evidence fields before using any record operationally.
