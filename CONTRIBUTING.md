# Contributing to DocuMint

First off, thanks for taking the time to contribute! 🎉

The following is a set of guidelines for contributing to DocuMint. These are mostly guidelines, not rules. Use your best judgment, and feel free to propose changes to this document in a pull request.

## Getting Started

1.  **Fork the repo** on GitHub.
2.  **Clone** the project to your own machine.
3.  **Create a Virtual Environment**:
    ```bash
    python -m venv venv
    source venv/bin/activate  # or .\venv\Scripts\activate on Windows
    ```
4.  **Install Dependencies**:
    ```bash
    pip install .  # Installs the package in editable mode
    pip install -r requirements.txt # Fallback
    ```

## Development Workflow

1.  Create your feature branch (`git checkout -b feature/amazing_feature`).
2.  Commit your changes (`git commit -m 'Add some amazing feature'`).
3.  Push to the branch (`git push origin feature/amazing_feature`).
4.  Open a Pull Request.

## Code Style

*   We follow PEP 8.
*   Please use type hints (`typing` module) for all new functions.
*   Add comments for complex logic.

## Running Tests

(Coming Soon) - We are implementing `pytest`. Please ensure any new code does not break existing functionality by performing a manual dry-run.
