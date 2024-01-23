const d3 = require('d3')

const config = require('../../config')
const { addPdfCoverTitle } = require('../pdfPage')
const featureToggles = config().featureToggles

function renderBanner(renderFullRadar) {
  if (featureToggles.UIRefresh2022) {
    const documentTitle = document.title[0].toUpperCase() + document.title.slice(1)

    document.title = documentTitle
    d3.select('.hero-banner__wrapper')
      .append('p')
      .classed('hero-banner__subtitle-text', true)
      .text("Dit is de Tech Radar van het Kadaster Emerging Technology Center. Hierin zijn nieuwe technologieën, ontwikkelingen en innovaties ingedeeld in hoe het ETC verwacht dat deze het Kadaster raken en/of relevant zijn. Heb je vragen, opmerkingen of suggesties? Stuur ons dan een mail naar etc@kadaster.nl")
      .append('p')
      .classed('hero-banner__subtitle-text', true)
      .text("Laatst bijgewerkt op 30 maart 2023")
    d3.select('.hero-banner__title-text').on('click', renderFullRadar)

    addPdfCoverTitle(documentTitle)
  } else {
    const header = d3.select('body').insert('header', '#radar')
    header
      .append('div')
      .attr('class', 'radar-title')
      .append('div')
      .attr('class', 'radar-title__text')
      .append('h1')
      .text(document.title)
      .style('cursor', 'pointer')
      .on('click', renderFullRadar)

    header
      .select('.radar-title')
      .append('div')
      .attr('class', 'radar-title__logo')
      .html('<a href="https://www.thoughtworks.com"> <img src="images/logo.png" /> </a>')
  }
}

module.exports = {
  renderBanner,
}
