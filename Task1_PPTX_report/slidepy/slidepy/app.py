import json
import collections 
import collections.abc
from typing import TypedDict
import typeguard
from pptx import Presentation
import logging

from slidepy.cli import CommandlineInterface, command, default_command
from slidepy.slides import Slide

class App (CommandlineInterface):
    """SlidePy

    This application is used to generate a report in pptx format based on a configuration file.
    """
    def __init__(self):
        # Init commandline interface
        super(App, self).__init__(prog='slidepy')

        # Configure logger
        self._configure_logger()

    def _configure_logger(self):
        # Create logger
        self.logger = logging.getLogger("slidepy")
        self.logger.setLevel(logging.INFO)

        # Create a file handler for logging to a file
        file_handler = logging.FileHandler("slidepy.log")
        file_handler.setLevel(logging.INFO)

        # Create a stream handler for logging to console
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(logging.INFO)

        # Create a formatter and add it to the handlers
        formatter = logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s", "%Y-%m-%d %H:%M:%S"
        )
        file_handler.setFormatter(formatter)
        stream_handler.setFormatter(formatter)

        # Add the handlers to the logger
        self.logger.addHandler(file_handler)
        self.logger.addHandler(stream_handler)


    @command
    def generate(self, config, output):
        """Report generation

        Use this subcommand to generate a report based on a configuration file.

        Args:
            config (str): Path to the configuration file
            output (str): Save the generated report into this file
        """
        # Initializing presentation
        presentation = Presentation()

        # Creating Config type for typechecking
        class Config(TypedDict):
            presentation:list[dict]

        # Loading config file from JSON
        try:
            with open(config, 'r') as config_file:
                data = json.load(config_file)
                typeguard.check_type(data, Config)
        except OSError:
            msg = f'Couldn\'t load "{config}"! Please check if it\'s a valid path.'
            self.logger.error(msg)
            return
        except json.JSONDecodeError as e:
            msg = f'Config file "{config}" is not a valid JSON file! {str(e)}'
            self.logger.error(msg)
            return
        except typeguard.TypeCheckError as e:
            msg = f'Config file "{config}" is not in a valid format! {str(e)}'
            self.logger.error(msg)
            return
        
        self.logger.info(f'Configuration file loaded: "{config}"')
        
        # Iterate through the presentation tag in the JSON and creating the
        # appropriate Slides
        for i, slide_properties in enumerate(data["presentation"]):
            try:
                Slide.from_dict(presentation, slide_properties)
            except typeguard.TypeCheckError as e:
                msg = f'Slide number {i+1} is not in a valid format! {str(e)}'
                self.logger.error(msg)
                return
            except ValueError as e:
                msg = f'Slide number {i+1} is not in a valid format! {str(e)}'
                self.logger.error(msg)
                return
            except KeyError as e:
                msg = f'Slide number {i+1} is missing a property! {str(e)}'
                self.logger.error(msg)
                return
            except OSError as e:
                msg = f'Slide number {i+1} has a file dependency which could not be loaded! {str(e)}'
                self.logger.error(msg)
                return
            
            self.logger.info(f'Slide number {i+1} generated')

        # Saving the presentation
        try:
            presentation.save(output)
        except OSError as e:
            msg = f'Couldn\'t save presentation to "{output}"! {str(e)}'
            self.logger.error(msg)
            return
        
        self.logger.info(f'Presentation saved: "{output}"')

    @default_command
    def default(self):
        self._parser.print_help()
