[![Prettier-Standard Style Guide](https://img.shields.io/badge/code_style-standard-brightgreen.svg)](https://github.com/sheerun/prettier-standard)

A fork of ThoughtWork's [build-your-own-radar](https://github.com/thoughtworks/build-your-own-radar) project.

## Demo

This radar is compiled and hosted on GitHub Pages. The `content/data.json` is plugged in, which leads to [this visualization](https://kadaster-labs.github.io/tech-radar/?documentId=https%3A%2F%2Fraw.githubusercontent.com%2Fkadaster-labs%2Fbuild-your-own-radar%2Fmaster%2Fcontent%2Fdata.json). This visualization is embedded on https://labs.kadaster.nl/tech-radar/.

## How To Use

### Using JSON data

We compile the list of technologies in a separate [Excel document](https://hetkadaster.sharepoint.com/sites/gd-boi/b08af/Content/2020-013-TechRadar/2024Q1-TechRadar.xlsx?d=w3a5cd46c7220490daa04478e7911d357&csf=1&web=1&e=VVexwD). This allows for easy collaboration between team members. Previously, we made CSV export from this document, but it has some complications with newlines and delimiters as the company uses the Dutch language pack.

Another, more convenient, way to provide the data is using a JSON array. The linked Excel document contains a "Office Scripts" script which outputs the Content table as JSON and prints it to console.log. This script can be used to easily export the data from Excel to this repo.
We can enter a publicly accessible URL (not behind any authentication) of this JSON file which will then be rendered. This file is hosted on GitHub, within this repo.

An example:

```json
[
  {
    "name": "Computer Vision",
    "ring": "Explore",
    "quadrant": "AI en Datascience",
    "isNew": "FALSE",
    "description": "Computer vision is een veld binnen AI ..."
  },
  {
    "name": "Generative AI",
    "ring": "Assess",
    "quadrant": "AI en Datascience",
    "isNew": "TRUE",
    "description": "Generatieve AI gaat over het genereren van ..."
  },
  {
    "name": "Event Stream Processing",
    "ring": "assess",
    "quadrant": "Moderne IT",
    "isNew": "FALSE",
    "description": "Data is steeds minder statisch en steeds meer ..."
  },
  {
    "name": "Self-Sovereign Identity",
    "ring": "Monitor",
    "quadrant": "Privacybescherming",
    "isNew": "FALSE",
    "description": "Self-sovereign identity (SSI) is een concept ..."
  }
]
```

We currently make is of the following quadrants. Please note this is subject to change:
- AI en Datascience
- Moderne IT
- Privacybescherming
- Federatieve Datadeling


### Building the radar

The develop this locally, we recommend Node 18. To install the dependencies:

```console
npm ci
```

In order to view the project in action, it needs to be built and served as web assets.

```console
npm run build:prod
npx serve dist/
```

The build artifacts can then be accessed in the browser at http://localhost:3000


### Hosting
We use GitHub Pages for hosting the compiled project. Every time a commit is made on the `main` branch, a GitHub Action workflow does the same steps as "Building the radar" and publishes the results as a GitHub Page.

The resulting visualization can be embedded on different websites. We use an iframe for this.

```html
<iframe class="tech-radar"
  src="https://kadaster-labs.github.io/tech-radar/?documentId=https%3A%2F%2Fraw.githubusercontent.com%2Fkadaster-labs%2Fbuild-your-own-radar%2Fmaster%2Fcontent%2Fdata.json"
  frameborder="0" height="1050px" width="100%">
</iframe>
```


### Contributing
For information on the e2e tests and usage of the Docker image, please consult the original build-your-own-radar repo.
