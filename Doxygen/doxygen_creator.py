# This Python file uses the following encoding: utf-8
"""
*****************************************************************************
 @file    doxygen_creator.py
 @brief   create doxygen documentation
*****************************************************************************
"""

import sys
import os
import subprocess
import webbrowser
import time
from threading import Thread
from doxygen import ConfigParser

# user include
import requests
sys.path.append('../')
import Source.wjh_sv as sv 

B_USE_OWN_STYLE = True
B_SIDEBAR_ONLY = True

S_DOXYGEN_PATH = "doxygen.exe" # required: add doxygen bin path to file path in system variables
S_DEFAULT_OUTPUT_FOLDER = "Output_Doxygen"
L_DEFAULT_FILE_PATTERN = ["*.py", "*.bat", "*.md"]
S_PUBLISHER = "Timo Unger"
S_MAIN_FOLDER_FOLDER = "../"
S_PLANTUML_PATH = "./" # need plantuml.jar in this folder
S_PLANTUML_JAR_URL = "https://github.com/plantuml/plantuml/releases/download/v1.2022.8/plantuml-1.2022.8.jar"
S_PLANTUML_JAR_NAME = "plantuml.jar"
S_WARNING_FILE_PREFIX = "Doxygen_warnings_"
S_WARNING_FILE_SUFFIX = ".log"
S_INDEX_FILE = "html/index.html"

YES = "YES"
NO = "NO"
WARNING_FAIL = "FAIL_ON_WARNINGS"

S_GITHUB_CORNER_FIRST = "<a href="
S_GITHUB_CORNER_LAST = """ class="github-corner" aria-label="View source on GitHub"><svg width="80" height="80" viewBox="0 0 250 250" style="fill:#151513; color:#fff; position: absolute; top: 0; border: 0; right: 0;" aria-hidden="true"><path d="M0,0 L115,115 L130,115 L142,142 L250,250 L250,0 Z"></path><path d="M128.3,109.0 C113.8,99.7 119.0,89.6 119.0,89.6 C122.0,82.7 120.5,78.6 120.5,78.6 C119.2,72.0 123.4,76.3 123.4,76.3 C127.3,80.9 125.5,87.3 125.5,87.3 C122.9,97.6 130.6,101.9 134.4,103.2" fill="currentColor" style="transform-origin: 130px 106px;" class="octo-arm"></path><path d="M115.0,115.0 C114.9,115.1 118.7,116.5 119.8,115.4 L133.7,101.6 C136.9,99.2 139.9,98.4 142.2,98.6 C133.8,88.0 127.5,74.4 143.8,58.0 C148.5,53.4 154.0,51.2 159.7,51.0 C160.3,49.4 163.2,43.6 171.4,40.1 C171.4,40.1 176.1,42.5 178.8,56.2 C183.1,58.6 187.2,61.8 190.9,65.4 C194.5,69.0 197.7,73.2 200.1,77.6 C213.8,80.2 216.3,84.9 216.3,84.9 C212.7,93.1 206.9,96.0 205.4,96.6 C205.1,102.4 203.0,107.8 198.3,112.5 C181.9,128.9 168.3,122.5 157.7,114.1 C157.9,116.9 156.7,120.9 152.7,124.9 L141.0,136.5 C139.8,137.7 141.6,141.9 141.8,141.8 Z" fill="currentColor" class="octo-body"></path></svg></a><style>.github-corner:hover .octo-arm{animation:octocat-wave 560ms ease-in-out}@keyframes octocat-wave{0%,100%{transform:rotate(0)}20%,60%{transform:rotate(-25deg)}40%,80%{transform:rotate(10deg)}}@media (max-width:500px){.github-corner:hover .octo-arm{animation:none}.github-corner .octo-arm{animation:octocat-wave 560ms ease-in-out}}</style>"""

# user defines
D_DOXYGEN_USER_SETTINGS = {
    'PROJECT_NAME' : sv.S_WJHSV_APPLICATION_NAME,
    'PROJECT_NUMBER' : f"v{sv.S_VERSION}",
    'PROJECT_BRIEF' : sv.S_WJHSV_DESCRIPTION,
    'PROJECT_LOGO' : f"{S_MAIN_FOLDER_FOLDER}{sv.S_ICON_RESOURCE_PATH}",
    'INPUT'        : S_MAIN_FOLDER_FOLDER,
    'USE_MDFILE_AS_MAINPAGE' : f"{S_MAIN_FOLDER_FOLDER}README.md"
}

S_REPO_LINK = "https://github.com/timounger/WJH-SV"


class OpenNotepad(Thread):
    def __init__(self, s_file):
        Thread.__init__(self)
        self.s_file = s_file
        self.start()
    def run(self):
        time.sleep(1)
        with subprocess.Popen(["notepad.exe", self.s_file]):
            pass

class DoxygenCreator():
    """!
    @brief Class to generate Doxygen documentation for and code documentation with uniform settings and style
    @param d_user_settings : induvidual settings so set in Doxyfile
    @param s_webside : webside link for GitHub corner
    """
    def __init__(self, d_user_settings = [], s_webside = None):
        self.d_user_settings = d_user_settings
        self.s_webside = s_webside

        if ('PROJECT_NAME' in self.d_user_settings):
            self.s_project_name = self.d_user_settings['PROJECT_NAME']
        else:
            self.s_project_name = ""

        if ('OUTPUT_DIRECTORY' in self.d_user_settings) and (self.d_user_settings['OUTPUT_DIRECTORY'] != ""):
            self.s_output_dir = f"{self.d_user_settings['OUTPUT_DIRECTORY']}/"
        else:
            self.s_output_dir = f"{S_DEFAULT_OUTPUT_FOLDER}/"
        self.s_doxyfile_name = f"{self.s_project_name}.Doxyfile"
        self.s_warning_name = f"{self.s_output_dir}{S_WARNING_FILE_PREFIX}{self.s_project_name}{S_WARNING_FILE_SUFFIX}"

    def create_default_doxyfile(self):
        """!
        @brief Create default doxyfile
        """
        if not os.path.exists(self.s_output_dir):
            os.makedirs(self.s_output_dir)
        subprocess.call([S_DOXYGEN_PATH, "-g", self.s_doxyfile_name])

    def edit_select_doxyfile_settings(self):
        """!
        @brief Edit settings in doxyfile
        """
        # load the default doxygen template file
        config_parser = ConfigParser()
        configuration = config_parser.load_configuration(self.s_doxyfile_name)

        # Clear lists at first
        configuration['INPUT'] = []
        configuration['FILE_PATTERNS'] = []
        configuration['EXCLUDE_PATTERNS'] = self.s_output_dir #ignore output folder

        # write user settings
        for key, value in self.d_user_settings.items():
            if isinstance(value, list):
                for entry in value:
                    configuration[key].append(entry)
            else:
                configuration[key] = value
        
        configuration['OUTPUT_DIRECTORY'] = self.s_output_dir # to add no empty out directory
        
        if ('PROJECT_BRIEF' not in self.d_user_settings) and (self.s_project_name != ""): # use project name as default if nothing defined
            configuration['PROJECT_BRIEF'] = f"{self.s_project_name}-Documentation"

        # other settings
        configuration['FULL_PATH_NAMES'] = NO
        configuration['OPTIMIZE_OUTPUT_JAVA'] = YES
        configuration['OPTIMIZE_OUTPUT_JAVA'] = YES
        configuration['EXTRACT_ALL'] = YES
        configuration['SORT_BY_SCOPE_NAME'] = YES
        configuration['WARN_NO_PARAMDOC'] = YES
        configuration['WARN_AS_ERROR'] = WARNING_FAIL
        if 'INPUT' not in self.d_user_settings:
            configuration['INPUT'].append('.')
        if 'FILE_PATTERNS' not in self.d_user_settings:
            for file in L_DEFAULT_FILE_PATTERN:
                configuration['FILE_PATTERNS'].append(file)
        configuration['RECURSIVE'] = YES
        configuration['SOURCE_BROWSER'] = YES
        configuration['INLINE_SOURCES'] = YES
        configuration['HTML_TIMESTAMP'] = YES
        configuration['HTML_DYNAMIC_SECTIONS'] = YES
        configuration['DOCSET_PUBLISHER_NAME'] = S_PUBLISHER
        configuration['MATHJAX_RELPATH'] = "https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.5/" # @todo prüfen ob doxygen es bereits korrekt macht
        configuration['GENERATE_LATEX'] = NO
        configuration['LATEX_CMD_NAME'] = 'latex'
        configuration['HAVE_DOT'] = YES
        configuration['UML_LOOK'] = YES
        configuration['DOT_IMAGE_FORMAT'] = "svg"
        configuration['INTERACTIVE_SVG'] = YES
        configuration['PLANTUML_JAR_PATH'] = S_PLANTUML_PATH
        configuration['PLANTUML_INCLUDE_PATH'] = S_PLANTUML_PATH
        configuration['DOT_MULTI_TARGETS'] = YES
        
        configuration['INTERNAL_DOCS'] = YES
        configuration['JAVADOC_AUTOBRIEF'] = YES
        configuration['HIDE_SCOPE_NAMES'] = YES
        configuration['SORT_BRIEF_DOCS'] = YES
        configuration['ABBREVIATE_BRIEF'] = ""
        #configuration['DOT_TRANSPARENT'] = YES
        configuration['WARN_LOGFILE'] = self.s_warning_name

        # required settings for the style template (do not change)
        configuration['GENERATE_TREEVIEW'] = YES
        configuration['DISABLE_INDEX'] = NO
        configuration['FULL_SIDEBAR'] = NO
        configuration['HTML_HEADER'] = "header.html"
        if B_USE_OWN_STYLE:
            if B_SIDEBAR_ONLY or (self.s_webside is not None):
                configuration['HTML_EXTRA_STYLESHEET'] = []
                configuration['HTML_EXTRA_STYLESHEET'].append("doxygen-awesome.css")
                configuration['HTML_EXTRA_STYLESHEET'].append("doxygen-awesome-sidebar-only.css")
                configuration['HTML_EXTRA_STYLESHEET'].append("doxygen-awesome-sidebar-only-darkmode-toggle.css")
            else:
                configuration['HTML_EXTRA_STYLESHEET'] = "doxygen-awesome.css"
            configuration['HTML_EXTRA_FILES'] = []
            configuration['HTML_EXTRA_FILES'].append("doxygen-awesome-darkmode-toggle.js")
            configuration['HTML_EXTRA_FILES'].append("doxygen-awesome-fragment-copy-button.js")
            configuration['HTML_EXTRA_FILES'].append("doxygen-awesome-paragraph-link.js")
            configuration['HTML_EXTRA_FILES'].append("doxygen-awesome-interactive-toc.js")
            configuration['HTML_COLORSTYLE_HUE'] = 209
            configuration['HTML_COLORSTYLE_SAT'] = 255
            configuration['HTML_COLORSTYLE_GAMMA'] = 113

        # store the configuration in doxyfile
        config_parser.store_configuration(configuration, self.s_doxyfile_name)

    def download_plantuml_jar(self):
        """!
        @brief Generate Doxygen output depend on existing doxyfile
        """
        if not os.path.exists(S_PLANTUML_JAR_NAME):
            print(f"Download {S_PLANTUML_JAR_NAME} ...")
            try:
                response = requests.get(S_PLANTUML_JAR_URL)
                open(S_PLANTUML_JAR_NAME, "wb").write(response.content) # download plantuml.jar
            except Exception:
                print(f"Can not downlaad {S_PLANTUML_JAR_NAME}! Generate documentation without PlantUml Graph Support.")
        else:
            print(f"{S_PLANTUML_JAR_NAME} already exist!")

    def check_doxygen_warnings(self, b_open_warning_file = True):
        """!
        @brief Check for doxygen Warnings
        @param b_open_warning_file : [True] open warning file; [False] only check for warnings
        """
        b_warnings = False
        
        if os.path.exists(self.s_warning_name):
            with open(self.s_warning_name) as file:
                lines = file.readlines()
                if len(lines) != 0:
                    b_warnings = True
                    print("Doxygen Warnings found!")

        if b_open_warning_file and b_warnings:
            OpenNotepad(self.s_warning_name)

        return b_warnings

    def generate_doxygen_output(self, b_open_doxygen_output = True):
        """!
        @brief Generate Doxygen output depend on existing doxyfile
        @param b_open_doxygen_output : [True] open output in browser; [False] only generate output
        """
        self.download_plantuml_jar()
        subprocess.call([S_DOXYGEN_PATH, self.s_doxyfile_name])

        if b_open_doxygen_output:
            # open doxygen output
            filename = f"file:///{os.getcwd()}/{self.s_output_dir}{S_INDEX_FILE}"
            webbrowser.open_new_tab(filename)

    def add_github_corner(self):
        """!
        @brief Add Github corner
        """
        if self.s_webside is not None:
            s_corner_text = S_GITHUB_CORNER_FIRST + self.s_webside+ S_GITHUB_CORNER_LAST
            s_folder = f"{self.s_output_dir}html/"
            for file in os.listdir(s_folder):
                if file.endswith(".html"):
                    with open(f"{s_folder}{file}", "a") as file:
                        file.write(s_corner_text)

    def run_doxygen(self, b_open_doxygen_output = True):
        """!
        @brief Generate Doxyfile and Doxygen output depend on doxyfile settings
        @param b_open_doxygen_output : [True] open output in browser; [False] only generate output
        """
        self.create_default_doxyfile()
        self.edit_select_doxyfile_settings()
        self.generate_doxygen_output(b_open_doxygen_output)
        self.add_github_corner()
        b_warnings = self.check_doxygen_warnings(b_open_doxygen_output)
        return b_warnings

if __name__ == "__main__":
    doxygen_creator = DoxygenCreator(D_DOXYGEN_USER_SETTINGS, S_REPO_LINK)
    sys.exit(doxygen_creator.run_doxygen())
