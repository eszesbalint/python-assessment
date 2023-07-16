## Installation Guide

To install the `slidepy` application, follow these steps:

1. Make sure you have Python 3.10 or higher installed on your system.

2. Clone the GitHub repository:

   ```
   git clone https://github.com/eszesbalint/python-assessment.git
   ```

3. Navigate to the `python-assessment/Task1_PPTX_report/slidepy` directory:

   ```
   cd python-assessment/Task1_PPTX_report/slidepy
   ```

4. Install the required dependencies by running the following command:

   ```
   pip install .
   ```

   This will install the necessary packages: `docstring-parser`, `numpy`, `typeguard`, `matplotlib`, and `python-pptx`.

5. The installation is now complete.

## Usage Guide

The `slidepy` application is used to generate a report in pptx format based on a configuration file. Follow the steps below to use the application:

1. Create a configuration file in JSON format. The configuration file should define the slides of the presentation. See the sample configuration file for reference.

2. Run the application using the following command:

   ```
   python -m slidepy generate --config <path_to_config_file> --output <output_file>
   ```

   Replace `<path_to_config_file>` with the path to your configuration file and `<output_file>` with the desired name and path for the generated report.

   For example:

   ```
   python -m slidepy generate --config config.json --output report.pptx
   ```

3. The application will read the configuration file, generate the report slides, and save the output file in pptx format.

4. If any errors occur during the generation process, error messages will be displayed in the console indicating the specific issues with the configuration file or slide properties.

   - If the configuration file path is invalid, an error message will be displayed.
   - If the configuration file is not a valid JSON file, an error message will be displayed.
   - If the configuration file does not adhere to the expected format, an error message will be displayed.
   - If a slide is not in a valid format or has missing properties, an error message will be displayed.
   - If a slide has a file dependency that could not be loaded, an error message will be displayed.

5. Once the generation process is complete, the application will save the generated report in the specified output file.

   - If the output file path is invalid or there are issues with saving the presentation, an error message will be displayed.

6. The generated report is now ready for use.

Note: You can also run the application without specifying the `generate` command, which will display the help text and available options:

```
python -m slidepy 
```

This will display the available commands, their descriptions, and the required arguments.