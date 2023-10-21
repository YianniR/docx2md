# DOCX to Markdown Converter

Convert your DOCX files to Markdown format with ease using this command-line tool. This tool simplifies the process of converting DOCX files to Markdown, making it suitable for a wide range of use cases, from content creators to technical writers.

## Table of Contents

* [Features](#features)
* [Getting Started](#getting-started)
  * [Prerequisites](#prerequisites)
  * [Installation](#installation)
* [Usage](#usage)
* [Options](#options)
* [Examples](#examples)
* [Contributing](#contributing)
* [License](#license)

## Features

* **Effortless Conversion:** Convert DOCX files to Markdown format quickly and effortlessly.
* **Media Extraction:** Extract embedded media (images, etc.) from DOCX files during conversion.
* **Recursive Search:** Option to search for DOCX files in subfolders for batch processing.
* **Progress Tracking:** A built-in progress bar provides real-time feedback during conversion.

## Getting Started

### Prerequisites

Before using the DOCX to Markdown Converter, ensure you have the following prerequisites:

* [Pandoc](https://pandoc.org/): This tool relies on Pandoc for the actual file conversion. You can download and install Pandoc from the [official website](https://pandoc.org/installing.html) or use your system's package manager.
* [Python](https://www.python.org/downloads/): The converter is a Python script, so you need Python installed on your system. You can download Python from the official website.

### Installation

1. Clone the repository to your local machine:
   <pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>

* <pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">git clone https://github.com/yourusername/docx-to-markdown-converter.git
  </code></div></div></pre>
* Change to the project directory:
  <pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>
* <pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">cd docx-to-markdown-converter
  </code></div></div></pre>
* Install the required Python packages using `pip`:
  <pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>

1. <pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">pip install -r requirements.txt
   </code></div></div></pre>

## Usage

To convert DOCX files to Markdown, use the following command:

<pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>

<pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">python convert.py folder [-r]
</code></div></div></pre>

* `folder`: The path to the folder containing DOCX files.
* `-r`, `--recursive`: (Optional) Search for DOCX files in subfolders.

## Options

* `folder`: The folder path to search for DOCX files.
* `-r`, `--recursive`: Search for DOCX files in subfolders.

## Examples

1. Convert all DOCX files in the current folder:
   <pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>

* <pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">python convert.py .
  </code></div></div></pre>
* Convert all DOCX files in a specific folder:
  <pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>
* <pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">python convert.py /path/to/your/folder
  </code></div></div></pre>
* Convert all DOCX files in a folder and its subfolders:
  <pre><div class="bg-black rounded-md mb-4"><div class="flex items-center relative text-gray-200 bg-gray-800 gizmo:dark:bg-token-surface-primary px-4 py-2 text-xs font-sans justify-between rounded-t-md"><span>bash</span></div></div></pre>

1. <pre><div class="bg-black rounded-md mb-4"><div class="p-4 overflow-y-auto"><code class="!whitespace-pre hljs language-bash">python convert.py /path/to/your/folder -r
   </code></div></div></pre>

## Contributing

Contributions to this project are welcome! If you have ideas for improvements or find any issues, please submit them on the [GitHub issue tracker]().

## License

This project is licensed under the MIT License. See the [LICENSE]() file for details.
