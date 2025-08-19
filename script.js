function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tristan script")
    .addItem("Analyse des declis", "ShowOptionsForm")
    .addItem("Remplir liste (active colonne)", "get_distinct_value")
    .addToUi();
}

function ShowOptionsForm() {
  const html = HtmlService.createHtmlOutputFromFile("modal")
    .setWidth(300)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Options d'analyse");
}

function CopyColumnProduct(ss) {
  let data = ss.getRange(1, 1, ss.getLastRow(), 1).getValues();

  ss.insertColumnsAfter(1, 1);
  ss.getRange(1,2, data.length, 1).setValues(data);
}

function get_color_list() {
  let list = ["vert","jaune", "bleu", "turquoise", "rouge", "rouge et noir", "vert", "noir", "bleu teal", "rose", "noir", "pistache", "vert menthe", "rouge", "bleu persan", "vert herbe", "corail",
             "blue", "white", "blue floral reversible", "butter yellow stripes", "floral green reversible", "floral pink reversible", "floral sage reversible", "green stripes palm trees  reversible",
              "seashell reversible", "pink stripes", "red stripes", "yellow", "honeydew", "mint/orange", "sage green", "floral print", "airy blue", "dots", "buttercream floral",
             "ekru", "beige", "terracotta", "vert sapin", "sauge", "bleu roi", "pistache", "coffee", "orange", "jaune", "lila", "baby pink", "mint", "peach", "ekru + pistache",
              "ekru + terrracotta", "ekru + beige", "ekru + sauge", "ekru + vert sapin", "ekru + coffee", "ekru + lila", "ekru + baby blue", "ekru + baby pink", "ekru + mint",
              "ekru + peach", "peach + pink", "pink + lila", "mint + blue", "beige + terracotta", "pistache + sapin", "noir", "matcha", "baby blue", "fuschia", "ekru + orange",
              "ekru + terracotta", "vert", "blanc", "terrracotta", "ekru + coffee + terracotta", "candy pink", "black", "leopard", "coral", "black leopard", "fiery red", "cream/black",
              "pine green", "papaya", "apricot", "matcha", "warm nude", "electric lilac", "cherry/coral", "cherry", "warm nude/candy pink", "pink", "cream", "powder pink/fiery red",
              "light yellow", "peach", "turquoise/peach", "powder pink", "coral/cherry", "turquoise/coral", "warm nude/coral/turquoise", "powder pink/light yellow", "light blue/fiery red",
              "candy pink/electric blue", "mocha/candy pink", "turquoise", "kelly green", "lavender", "mauve", "peach/turquoise", "nude", "warm nude/lilac", "neon pink", "off white", "navy blue",
              "pale pink/burgundy", "apricot/neon pink", "electric blue", "dusty rose", "sky blue", "turquoise blue", "cobalt blue", "olive", "faded blue", "grey"];
  return (list);
}

function get_size_list() {
  let list = ["200 ml", "47", "2 oz", "12cm", "22cm", "19cm", "20cm", "27cm", "17.5cm", "23cm", "taille s (9cm)", "taille m (12cm)", "taille l (20cm)", "50ml", "15ml", "100ml", "200ml", "100g"];
  return (list);
}

function get_material_list() {
  let list = ["cuir", "or", "argent"];
  return (list);
}

function get_taste_list() {
  let list = ["vanille", "pistache", "brownie cacahuètes", "choco noisette", "cacao noisette-praliné", "cookie dough", "cacao-noisette", "cacao-cacahuète", "praliné", "cacahuète", "chocolat", "miel", "beurre de cacahuète & cacao", "cacao noisette", "cacao"];
  return (list);
}

function get_fragrance_list() {
  let list = ["lavande", "poivre", "camomille", "amande", "herbe zen", "rose", "matcha", "vanille", "fleur d'oranger", "fleur de cerisier", "thé blanc", "tonka", "bois de oud", "cèdre", "figuier",
              "fleur de coton", "fleur de jasmin", "freesia", "gingembre", "grenade", "iris", "litchi", "magnolia", "musc", "osmanthus", "pistache", "pomelo", "santal", "thé matcha", "tiaré", "vétiver", "ylang", "yuzu", "cerisier"];
  return (list);
}

function get_jewelry_list() {
  let list = ["collier", "bague", "boucle d'oreille", "bracelet"];
  return (list);
}

function sortlist(list) {
  list.sort((a, b) => b.length - a.length);
  return list;
}

function decli_analysis(list, products) {
  let new_column = [];
  let new_name_product = [];

  list = sortlist(list);

  for (let i = 0; i < products.length; i++) {
    let row_product = products[i][0];
    let found = "";

    for (let j of list) {
      let regex = new RegExp("\\b" + j + "\\b", "i");
      if (regex.test(row_product)) {
        found = j;
        row_product = row_product.replace(regex, "");
        break;
      }
    }
    row_product = row_product.replace(/\s+/g, " ").trim();
    found       = found.charAt(0).toUpperCase() + found.slice(1).toLowerCase();

    new_name_product.push([row_product]);
    new_column.push([found]);
  }
  return {new_name_product, new_column};
}

function processSelections(decli_list) {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let all_lists = {
    "couleur" : get_color_list,
    "taille"  : get_size_list,
    "matiere" : get_material_list,
    "gout"    : get_taste_list,
    "parfum"  : get_fragrance_list,
    "bijoux"  : get_jewelry_list
    // "option"  : option_analysis
  };

  CopyColumnProduct(ss);
  let copy_data = ss.getRange(2, 2, ss.getLastRow() - 1, 1).getValues();

  for(var i = 0; i < decli_list.length; i++){
    var key = decli_list[i];
    if (all_lists[key]) {
      let list = all_lists[key]();
      let res = decli_analysis(list, copy_data);

      copy_data = res.new_name_product;
      ss.insertColumnsAfter(2, 1);
      ss.getRange(1,3).setValue(key.charAt(0).toUpperCase() + key.slice(1));
      ss.getRange(1, 3).setFontWeight("bold").setHorizontalAlignment("center");
      ss.getRange(2, 3, res.new_column.length, 1).setValues(res.new_column);
    }
  }
  ss.getRange(2, 2, copy_data.length, 1).setValues(copy_data);
}

///////////////////////// Script pour remplir les listes ////////////////////////////

function get_distinct_value() {
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let col = ss.getActiveCell().getColumn();
  let data = ss.getRange(2, col, ss.getLastRow() - 1, 1).getValues().flat();

  let unique = [...new Set(data.filter(String).map(v => v.trim().toLowerCase()))];

  let formated = unique.map(v => `"${v}"`).join(", ");

  ss.getRange(1, 26).setValue(formated);

  SpreadsheetApp.getUi().alert("Liste générée en Z1, prête à copier !");
}
