# gdocs-tests
Google docs quiz generation framework

Allows to generate randomized individual quiz variants from a single template.

## Installation

1. Create a new google docs document
2. Go to `Extensions` -> `Apps Script`
3. Paste the ./src contents into the GApp

## Usage 

The script will add a new menu item `Generate tests` to the google docs document.

## Syntax
The syntax is inteneded to be a close to a simple text as possible, so it's easier to copy tests from other sources.

Test entry:

1 Question
1.1 Answer 1
1.2 Answer 2
1.3 Answer 3
1.4 Answer 4

The first answer is the correct one by default.