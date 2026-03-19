# HO3 Forms Automation

A public-safe demo project for generating form combinations and organizing a document generation workflow.

## Overview

This repository shows how to:

* load a list of forms from a source file
* separate mandatory and optional forms
* apply mutual-exclusion rules
* generate valid form combinations
* export a combination matrix for review

This public version is intentionally simplified and sanitized for demonstration purposes.

## Why This Project Exists

This project is based on a real-world workflow pattern often used in document-heavy operations:

1. start with a source list of forms
2. apply combination rules
3. generate valid outputs
4. organize downstream document production

The goal of this repository is to demonstrate project structure, rule-based combination generation, and clean code organization in a way that is safe for public sharing.

## Public Demo Scope

This repository does **NOT** include:

* internal company file paths
* private Excel templates
* private PDF files
* production business data
* confidential underwriting logic

Instead, it includes a simplified public demo version of the workflow.

## Project Structure

```text
src/ho3_forms/       core application logic
scripts/             runnable scripts
config/              public-safe configuration files
data/sample/         sample input data
output/              generated outputs
tests/               test placeholders
```

## Quick Start

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the demo script:

```bash
python scripts/generate_combo_matrix.py
```

## Current Status

This is a public-safe portfolio version of the project.

It is designed to show:

* project organization
* Python scripting structure
* rule-driven combination generation
* safe separation of code, configuration, and data

## Future Improvements

Possible future additions include:

* configuration-driven rules
* test coverage
* CLI arguments
* document-generation workflow stubs
* sample reporting outputs

## License

This project is shared as a public demo for educational and portfolio purposes.
