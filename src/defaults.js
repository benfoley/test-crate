// Built-in defaults, from corpus-tools-dyirbal config.json / sample-data.json.
// Overridden by config.json / sample-data.json in the chosen folder if present.
// config.dataDir is intentionally ignored — the chosen folder is the data dir.

export const DEFAULT_CONFIG = {
  rootDataset: {
    "@id": "arcp://name,dyirbal-workshop",
    "@type": ["Dataset", "RepositoryCollection"],
    conformsTo: { "@id": "https://w3id.org/ldac/profile#Collection" },
    name: "Dyirbal and Neighbouring Languages Corpus",
    description:
      "A collection of digitised resources — dictionaries, wordlists, articles and field materials — relating to Dyirbal and neighbouring North Queensland languages.",
    datePublished: "2026-07-07",
    inLanguage: { "@id": "https://glottolog.org/resource/languoid/id/stan1293" },
    license: {
      "@id": "https://language-research-technology.github.io/qa/licenses/dyirbal/test/v1/",
      "@type": "ldac:DataReuseLicense",
      name: "Dyirbal Test Licence",
      description: "This is a placeholder licence",
      metadataIsPublic: true,
      allowTextIndex: true,
    },
    creator: { "@id": "#ben-foley", "@type": "Person", name: "Ben Foley" },
  },
  metadataLicence: {
    "@id": "https://creativecommons.org/licenses/by/4.0/",
    "@type": "ldac:DataReuseLicense",
    name: "Attribution 4.0 International (CC BY 4.0)",
    description:
      "You are free to: Share — copy and redistribute the material in any medium or format. Adapt — remix, transform, and build upon the material for any purpose, even commercially. This license is acceptable for Free Cultural Works. The licensor cannot revoke these freedoms as long as you follow the license terms.",
    metadataIsPublic: true,
    allowTextIndex: true,
  },
};

export const DEFAULT_SAMPLE_DATA = {
  people: [
    { "@id": "#ben-foley", "@type": "Person", name: "Ben Foley" },
    { "@id": "#tati-florez", "@type": "Person", name: "Tati Florez" },
    { "@id": "#des-crump", "@type": "Person", name: "Des Crump" },
  ],
  places: [
    { "@id": "#place-tully-river", "@type": "Place", name: "Tully River", description: "A river in Far North Queensland, Australia, flowing from the Atherton Tableland to the Coral Sea near Tully.", geo: { "@id": "#locality-tully-river" } },
    { "@id": "#place-brisbane", "@type": "Place", name: "Brisbane", description: "The capital city of Queensland, Australia.", geo: { "@id": "#locality-brisbane" } },
    { "@id": "#place-palm-island", "@type": "Place", name: "Palm Island", description: "An Aboriginal community on the Great Palm Island group, off the coast of Queensland near Townsville.", geo: { "@id": "#locality-palm-island" } },
  ],
  localities: [
    { "@id": "#locality-tully-river", "@type": "Geometry", asWKT: "POINT(145.9167 -17.9333)" },
    { "@id": "#locality-brisbane", "@type": "Geometry", asWKT: "POINT(153.0251 -27.4698)" },
    { "@id": "#locality-palm-island", "@type": "Geometry", asWKT: "POINT(146.5833 -18.75)" },
  ]
};

// Custom rdf:Property definitions added to the graph (as in the original index.js).
export const CUSTOM_PROPERTIES = [
  { "@id": "arcp://name,custom/terms#possibleDuplicate", "@type": "rdf:Property", name: "Possible Duplicate", description: "Filename of a possible duplicate." },
  { "@id": "arcp://name,custom/terms#participant", "@type": "rdf:Property", name: "Participant", description: "A participant associated with the file." },
  { "@id": "arcp://name,custom/terms#compiler", "@type": "rdf:Property", name: "Compiler", description: "The person who compiled the file." },
  { "@id": "arcp://name,custom/terms#austlangCode", "@type": "rdf:Property", name: "Austlang Code", description: "The AUSTLANG code for a language." },
  { "@id": "arcp://name,custom/terms#iso639-3", "@type": "rdf:Property", name: "ISO 639-3", description: "The ISO 639-3 code for a language." },
  { "@id": "arcp://name,custom/terms#glottologCode", "@type": "rdf:Property", name: "Glottolog Code", description: "The Glottolog code for a language." },
];
