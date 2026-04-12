# Overview

Azure DevOps has provided a menu for initiating a classic Release from a build, however this functionality was never implemented for YAML release pipelines.

This extension adds a `Next` tab to pipeline runs.

![Screenshot of Next Run tab](images/run-next-1.png)

# Instructions

1. Create two YAML pipelines: a build, and a release.
    - The release pipeline must reference the build pipeline as a [pipeline resource](https://learn.microsoft.com/en-us/azure/devops/pipelines/yaml-schema/resources-pipelines-pipeline?view=azure-pipelines).
2. Open the `Next` tab from a completed **Release** run.
    - If the YAML pipelines are valid, the list of *source pipelines* updates with the name of the build.
    - This step only needs to be performed once per build/release pair.
3. Open the `Next` tab from a completed **Build** run, then either:
    - Select a release and tap the **Run** button.
    - Double-click a release.
