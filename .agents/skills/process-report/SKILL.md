---
name: process-report
description: Process a Microsoft Publisher report file, convert it into markdown and store it in a knowledgebase.
---

When invoked, the user should specify a MS Publsher file that you are to ingest. If it's unclear which one, prompt the user.

Invoke this projects ./python/pub2md.py file against the input publisher file, store the output in the ./processed-reports directory of this project.
