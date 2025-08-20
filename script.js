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
              "pale pink/burgundy", "apricot/neon pink", "electric blue", "dusty rose", "sky blue", "turquoise blue", "cobalt blue", "olive", "faded blue", "grey", "beige", "marron", "noires marbrées",
              "rouge", "vertes", "blanches nacrées", "bleu océan", "bleu pastel", "bleues givrées", "bordeaux", "fuschias", "fuschias nacrées", "ivoires", "jaune pastel", "lilas", "marrons",
              "mocha", "noires", "nougat glacé", "olive", "oranges", "rose pâle", "roses", "taupe", "terracotta", "vert d'eau", "vert pastel", "vert sapin", "vieux rose", "violet",
              "bleu ciel", "corail", "emeraude", "mauve", "orange fluo", "rose pale", "dorées", "ecaille", "kaki", "léopard", "rose", "tacheté", "ivoire/vert d'eau", "violet/vert d'eau", "argentées", "blanches", "bleu roi", "jaune/pêche",
              "mermaid", "mint", "pailleté bronze", "pailleté noir", "pailleté rose", "pailleté vert", "pailleté violet", "pivoine", "tachetées", "vanille/rose", "vert/rose", "vert/violet", "vert/bleu", "bleu", "orange", "vert", "gris",
              "fuschia", "blanc", "doré", "noir", "coeur au choix", "ecru", "multi", "bronze", "terrazzo", "prune", "beige et blanc nacré", "blanc granité et noir onyx", "blanc ivoire", "blanc neige", "blanc, orange, bleu et vert", "bleu",
              "bleu clair / rouge / jaune", "bleu et blanc nacré", "bleu foncé", "bleu marine", "bleu nuit", "dégradé beige et marron", "dégradé bleu", "dégradé bleu gris kaki", "dégradé gris vert bouteille et beige", "dégradé marron volcanique",
              "dégradé rose poudre et beige", "gris anthracite", "gris bleuté", "gris orage", "gris lune", "jaune ocre", "kaki et bleu", "noir", "vert", "vert olive", "bois et noir", "bois vernis", "gris métal", "inox et noir mat", "inox mat et bois",
              "inox mat et noir mat", "noir mat", "inox", "inox mat", "noir brillant", "cuivré brillant", "doré mat", "gris clair", "blanc crème", "chocolat", "jaune moutarde", "orange", "rose pâle", "vert de gris", "vert pâle", "vert sapin", "beige",
              "gris", "marron", "blanc crème et vert de gris", "bleu nuit et gris anthracite","jaune moutarde et vert sapin", "orange et chocolat", "rose pâle et vert pâle", "beige et bois vernis", "ecru", "brigerton", "denim", "fushia", "glitter",
              "lilas", "vichy", "écaille", "rose", "rouge", "bleu klein", "bleu ciel", "fuchsia", "crème", "marron", "beige", "argile", "vert", "tomette", "olive", "argent", "orange", "jeans", "léopard", "noisette", "ciel", "noir", "noire", "seersucker",
              "ocre", "gris", "cérémonie", "bridgerton", "fleurs", "faded", "neon", "seaside", "cherry", "rose", "aqua", "berry", "tan", "palm", "lavender haze", "periwinkle", "sunshine", "hot pink", "caribbean checks", "coastal cowgirl", "azalea", "army",
              "charcoal", "algae", "earth tone", "caramel", "wavy checks navy", "wavy checks pine", "navy", "seafoam", "papaya", "emerald", "shave ice", "birch", "faded grey", "gold", "coconut milk", "youth", "aloha sunset neon", "calico crab charcoal",
              "calico crab faded army", "calico crab khaki", "calico crab salmon", "color block mango", "color block sand", "fleurs sunshine", "hazy daze bubblegum", "hazy daze swirl", "hazy daze turquoise", "palmitos seafoam", "sea abyss earth tone",
              "sea ripple tie dye", "sun daze charcoal", "sun daze yellow", "sunshine space honey", "sunshine space tahiti", "surfing cowboy bonanza blue", "surfing cowboy cactus", "triple scoop berry", "triple scoop blue moon", "triple scoop bubblegum",
              "triple scoop butter pecan", "triple scoop pistachio", "triple scoop velvet", "wavy checks avocado", "wavy checks banana", "wavy checks cocoa", "wavy flowers navy blue", "faded mustard yellow", "banana", "pine", "grilled cheese", "monday checkers",
              "sky blue", "raspberry", "mint chip", "pua party", "tropics patch", "adult", "floral", "black & white checkerboard", "checker moss", "black & white checker - strand", "dogtown mint", "turquoise venice", "olive", "sea glass", "clay", "mustard",
              "go with the flow desert", "go with the flow peach", "plankton green", "sun bleached", "patrick pink", "hibiscus", "retro daydream chartreuse", "retro daydream surfy 60s", "smiley khaki", "beach fossils", "raph red", "sewage grey", "youth black zipper hoodie",
              "cowabunga break up", "mutant checker turtle green", "static mayhem", "cloud", "almond", "khaki", "pesce", "pink and rust-orange", "sea moss", "daisy", "tahiti", "sand", "bonanza blue", "cactus", "coral", "fuchsia", "avocado", "cocoa", "lavender", "kiwi", "tan camo",
              "panama", "navy blue", "retro daydream floral", "seaside gingham", "autumn", "blush", "bubblegum pink", "mint", "seaweed", "soft sage", "oatmeal", "ivory", "cream", "dark brown", "soft pink", "blue", "vintage cream", "taupe", "emerald green",
              "olive", "black", "navy", "soft blue", "light wash", "soft grey", "pink", "mineral tan", "denim", "tan", "vintage sage", "vintage teal", "soft gry", "white", "stone", "soft peach", "medium wash", "light green", "beige", "sage", "green",
              "lavender", "gold", "teal", "olive green", "deep blue", "dark olive", "dark burgundy", "charcoal"];
  return (list);
}

function get_size_list() {
  let list = ["200 ml", "47", "2 oz", "12cm", "22cm", "19cm", "20cm", "27cm", "17.5cm", "23cm", "taille s (9cm)", "taille m (12cm)", "taille l (20cm)", "50ml", "15ml", "100ml", "200ml", "100g", "small - 2/3", "medium - 4/5", "toddler", "youth", "18/24", "1/3", "2/4", "3/5",
              "4/6", "5/7", "16", "6/8", "7/9", "14", "adult", "8/10", "extra small 2/3 years", "small 4/5 years", "medium 5/6 years", "extra large 10/11 years", "adult small - 12 years", "extra large - 8/9", "12", "extra small", "small", "medium", "large", "extra large",
              "36", "38", "40", "30", "34", "32", "extra small - 1yr", "small - 2yr", "medium - 5yr", "large - 6yr", "xs - 1 yr", "sm - 2 yr", "md - 5 yr", "5t", "24", "26", "28", "extra small - 2/3", "extra large -adult small", "4t", "6t", "9y", "22", "extra large - adult small",
              "xxl", "1t", "2t", "3t", "10y", "7y", "8y", "youth extra small - 2/3 years", "youth small - 4/5 years", "youth medium - 5/6 years", "youth large - 7/8 years", "youth extra large - 10/11", "adult small", "adult medium", "adult large", "adult extra large", "small - 4/5",
              "medium - 6/7", "large - 8/9", "extra large - 10/11", "medium - 5/6", "large - 7/8", "one size (adult + big kids)", "xx large", "large - 6/7", "9t", "small -2/3", "xxl - 12/14", "xxxl - 16/18", "extra small -2/3 years", "small - 4/5 years", "medium - 5/6 years",
              "large - 7/8 years", "extra large - 10/11 years", "xxxl - 16/18 - adult small", "adult small - 11/12 years", "adult medium - 13/14 years", "xxl - 10/11", "big kid", "extra small - 18 months"];
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
