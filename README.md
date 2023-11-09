# Corporate and Foundation Team MSWE Capstone

## Stakeholders

* Project Sponsor: Marianne Smith

* Software Developers: Tianyao Chen, Peiming Chen, Huikun Zheng, Xiaowen Sun

## Description

* Background/Current Situation/Problem/Opportunity: The team responsible for Corporate and Foundation Relations routinely invests a lot of effort in converting Pivot's CSV exports into user-friendly Word documents. This manual process not only hampers productivity but also leads to inconsistencies in the final document's look and feel.

* Overview of the desired impact: The goal is to create a specialized software solution for the team that will automate and simplify the document creation process. This will notably cut down on time spent on formatting, while also guaranteeing that the end products are consistently formatted and professional in appearance.

We call our software "AutoGranter".

## How to run the script

First time:

```sh
pip install virtualenv
virtualenv myenv
source myenv/bin/activate
pip install -r requirements.txt
# When you are done, deactivate
deactivate
```

Also add the file `apikey.txt` to the root directory of the project with the OpenAI API key inside.

After first use:

```sh
source myenv/bin/activate
# When you are done, deactivate
deactivate
```

## Tech Stack

1. Python: Core programming language for backend development and data manipulation.
2. pandas: Library for reading the CSV file.
3. python-docx: Library for generating Word documents.
4. Tkinter: To develop a simple, intuitive user interface..
5. GPT API: Utilized for natural language processing tasks, including auto-summarization.
