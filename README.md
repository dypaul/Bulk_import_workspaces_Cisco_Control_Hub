# Bulk_import_workspaces_Cisco_Control_Hub

# Your Project Name

A brief description of what your project does and why it's interesting.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Installation

Follow these steps to set up your development environment and install the project's dependencies.

1. **Clone the Repository:**

   Clone this GitHub repository to your local machine using the following command:

   ```bash
   git clone https://github.com/dypaul/Bulk_import_workspaces_Cisco_Control_Hub.git

2. **Navigate to the Project Directory:**
   
   Change your current directory to the project's directory:
   
   ```bash
   cd Bulk_import_workspaces_Cisco_Control_Hub

3. **Create a Virtual Environment (Optional but Recommended):**

   It's a good practice to create a virtual environment to isolate your project's dependencies. You can create one using Python's built-in venv module:
   
   ```bash
   python -m venv Bulk_import_workspaces_Cisco_Control_Hub
    ```
   Activate the virtual environment:

   On Windows:
    ```bash
   venv\Scripts\activate
    ```
    
   On macOS and Linux:
  
    ```bash
   source venv/bin/activate
    ```
4. **Install Dependencies:**

    Use pip to install the project's dependencies from the requirements.txt file:
   
    ```bash
    pip install -r requirements.txt
    ```
    This will install all the required packages and libraries.

## Usage
Our project requires you to configure it with your own API token for authentication. Follow these steps to configure the project:

1. **Obtain an API Token:**

   To use our project, you'll need to obtain an API token from [Webex Developper - Control Hub](https://developer.webex.com/docs/getting-started). This token is required for authentication and authorization.

2. **Replace the Token in the Code:**

   Open the project's source code (Bulk_import_workspaces.py) and locate the section where the API token needs to be provided. The variable named `access_token`.

3. **Insert Your API Token:**

   Replace the placeholder `YOUR_API_TOKEN_HERE` with your actual API token obtained from [Webex Developper - Control Hub](https://developer.webex.com/docs/getting-started).

4. **Save the Changes:**

   Save the file with the updated token.

5. **Execute the script:**
   Now that you've configured the project with your API token, you can execute it as follows:
   ```bash
   python Bulk_import_workspaces.py
   ```

**Note:** We recommend testing the script first on a Webex sandbox environment before using it in a production environment. You can create a sandbox account by following the instructions in the [Webex Developer Sandbox Guide](https://developer.webex.com/docs/developer-sandbox-guide).

## Contributing

   Code created with ChatGPT assistance

   We welcome contributions from the community to improve and enhance this project. Whether you want to fix a bug, add a new feature, or simply improve the documentation, your help is greatly appreciated.

## License

   This project is open-source and available under the [MIT License](LICENSE.md).

© Dylan PAUL (dypaul)
