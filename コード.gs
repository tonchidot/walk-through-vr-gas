function convertToNamedTable(data) {
  var header = data[0];
  var results = [];
  for (var i=1; i < data.length; i++) {
    var row = data[i];
    var rowObj = {};
    for (var j=0; j < row.length; j++) {
      rowObj[header[j]] = row[j];
    }
    results.push(rowObj);
  }
  return results;
}

function findArea(building, areaName) {
  for (var i=0; i < building.areas.length; i++) {
    if (building.areas[i].name == areaName) {
      return building.areas[i];
    }
  }
  return null;
}

function findSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var activeName = sheet.getName().replace(/(spot_images|route_images|areas|warp_spots|buildings|boards|annotations)/,"");
  var sheets = ss.getSheets();
  var results = { name: activeName };
  for (var i=0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name == "spot_images" + activeName) {
      results.spot_images = sheets[i];
    } else if (name == "route_images" + activeName) {
      results.route_images = sheets[i];
    } else if (name == "areas" + activeName) {
      results.areas = sheets[i];
    } else if (name == "warp_spots" + activeName) {
      results.warp_spots = sheets[i];
    } else if (name == "buildings" + activeName) {
      results.buildings = sheets[i];
    } else if (name == "boards" + activeName) {
      results.boards = sheets[i];
    } else if (name == "annotations" + activeName) {
      results.annotations = sheets[i];
    }
  }
  return results;
}

// name
function convertBuildings(building, sheet){
  if(sheet) {
    var data = convertToNamedTable(sheet.getDataRange().getValues());
    building.name = data[0].name
  }
  return building
}

// name | start_spot | start_area | image_path | image_width | image_height | image_px_per_meter
function convertAreas(building, sheet){
  var data = convertToNamedTable(sheet.getDataRange().getValues());
  for (var i=0; i < data.length; i++) {
    var row = data[i];
    var area = {
      name: row.name,
      start_spot: row.start_spot,
      layout: {
        image: {
          path: row.image_path,
          width: parseInt(row.image_width),
          height: parseInt(row.image_height),
          px_per_meter: parseFloat(row.image_px_per_meter)
        }
      },
      routes: [],
      spots: [],
      annotations: []
    };
    if (row.start_area === 'YES') {
      building.start_area = area.name;
    }
    building.areas.push(area);
  }
  return building;
}

// name | area_name | show_on_layout | default_rotation | image_name | image_path | image_rotation | position_x_px | position_z_px | position_x_m | position_z_m
function convertSpots(building, sheet) {
  var data = convertToNamedTable(sheet.getDataRange().getValues());
  for (var i=0; i < data.length; i++) {
    var row = data[i];
    var area = findArea(building, row.area_name);
    var spot = {
      name: row.name,
      show_on_layout: (row.show_on_layout === "YES"),
      default_rotation: {
        y: parseFloat(row.default_rotation) || 0.0
      },
      image: {
        path: row.image_path,
        rotation: {
          y: parseFloat(row.image_rotation)
        }
      },
      position: {
        x: parseFloat(row.position_x_m),
        z: parseFloat(row.position_z_m)
      }
    };
    area.spots.push(spot);
  }
  return building;
}

// from_spot_name | to_spot_name | area_name | image_name | image_path | image_rotation
function convertRoutes(building, sheet) {
  var data = convertToNamedTable(sheet.getDataRange().getValues());
  var routes = {};
  var names = [];
  Logger.log("# of routes: " + data.length);
  for (var i=0; i < data.length; i++) {
    var row = data[i];
    Logger.log("routes["+i+"]: " + Object.keys(row).map(function(key) { return key + ": " + row[key];}).join(", "));
    var name = [row.area_name, row.from_spot_name, row.to_spot_name].join("__");
    if (routes[name] === undefined) {
      routes[name] = {
        from: row.from_spot_name,
        to: row.to_spot_name,
        images: []
      }
      names.push({ name: name, area_name: row.area_name });
    }
    if(row.image_path !== "" && row.image_name !== "" && row.image_rotation !== "" ){
      routes[name].images.push({
        path: row.image_path,
        rotation: {
          y: parseFloat(row.image_rotation)
        }
      });
    }
  }
  for (var i=0; i < names.length; i++) {
    var route = names[i];
    var area = findArea(building, route.area_name);
    area.routes.push(routes[route.name]);
  }
  return building;
}

// src_area_name | src_spot_name | dst_area_name | dst_spot_name | direction | position_x_px | position_z_px | position_x_m | position_z_m
function convertWarpSpots(building, sheet) {
  if(sheet) { 
    var data = convertToNamedTable(sheet.getDataRange().getValues());
    for (var i=0; i < data.length; i++) {
      var row = data[i];
      var name = row.warp_spot_name === undefined ? '' : row.warp_spot_name;
      var position = { x: row.position_x_m, z: row.position_z_m };
      var source = { area_name: row.src_area_name, spot_name: row.src_spot_name };
      var destination = { area_name: row.dst_area_name, spot_name: row.dst_spot_name };
      var obj = { position: position, direction: row.direction, source: source, destination: destination };
      if (name !== '') {
        obj.name = name;
      }
      building.warp_spots.push(obj);
    }
  }
  return building;
}

// source_area_name | source_spot_name | destination_area_name | destination_spot_name | message
function convertBoards(building, sheet) {
  if(sheet) { 
    var data = convertToNamedTable(sheet.getDataRange().getValues());
    for (var i=0; i < data.length; i++) {
      var row = data[i];
      var source_name = row.source_area_name + '-' + row.source_spot_name;
      if (isWarp(building, row.destination_spot_name)) {
        var destination_name = row.destination_spot_name;
      } else {
        var destination_name = row.destination_area_name + '-' + row.destination_spot_name;
      }
      var message = row.message;
      building.boards.push({ source_name: source_name, destination_name: destination_name, message: message });
    }
  }
  return building;
}

// position_x | position_y | position_z | title | description | visible_spot_names
function convertAnnotations(building, sheet) {
  if(sheet) { 
    var data = convertToNamedTable(sheet.getDataRange().getValues());
    for (var i=0; i < data.length; i++) {
      var row = data[i];
      var area = findArea(building, row.area_name);
      var position = { x: row.position_x_m, y: row.position_y_m, z: row.position_z_m };
      var title = row.title;
      var description = row.description;
      var obj = { position: position, title: title, description: description };
      if(row.visible_spot_names !== '') {
        obj.visible_spot_names = row.visible_spot_names.split(',').map(function(spot_name){ return spot_name.trim(); });
      }
      area.annotations.push(obj);
    }
  }
  return building;
}

function generateJson(){
  SpreadsheetApp.flush();

  var res = findSheets();
  Logger.log(res['buildings'] ? "found "+res['buildings'].getName() : "buildings sheet not exist.");
  Logger.log("found "+res['spot_images'].getName());
  Logger.log("found "+res['route_images'].getName());
  Logger.log("found "+res['areas'].getName());
  Logger.log(res['warp_spots'] ? "found "+res['warp_spots'].getName() : "warp spots not exist.");
  Logger.log(res['boards'] ? "found "+res['boards'].getName() : "boards not exist.");
  Logger.log(res['annotations'] ? "found "+res['annotations'].getName() : "annotations not exist.");

  var building = {name: res.name, areas: [], warp_spots:　[], boards: [] };
  building = convertBuildings(building, res['buildings']);
  building = convertAreas(building, res['areas']);
  building = convertSpots(building, res['spot_images']);
  building = convertRoutes(building, res['route_images']);
  building = convertWarpSpots(building, res['warp_spots']);
  building = convertBoards(building, res['boards']);
  building = convertAnnotations(building, res['annotations']);
  // Logger.log(building.areas[0].spots[0]);
  // Logger.log(building.areas[0].routes[0]);
  // Logger.log(building.warp_spots);
  Logger.log('test');

  return building;
}

function showPreviewMsg(name, jsonUrl) {
  var previewUrl = "https://walk-through-preview.s3.amazonaws.com/preview/preview.html?json=" + encodeURIComponent(jsonUrl);
  var stagingUrl = "https://walk-through-preview.s3.amazonaws.com/preview/preview_staging.html?json=" + encodeURIComponent(jsonUrl);
  var productionUrl = "https://walk-through-preview.s3.amazonaws.com/preview/preview_production.html?json=" + encodeURIComponent(jsonUrl);
  var showSpotNameParam = '&showSpotName=true';
  var content = [
    '<div>',
    '<h3>' + name + 'のプレビュー</h3>',
    '<div><a target="_blank" href="' + previewUrl + '">プレビュー (preview環境) </a></div>',    
    '<div><a target="_blank" href="' + stagingUrl + '">プレビュー (staging環境) </a></div>',
    '<div><a target="_blank" href="' + productionUrl + '">プレビュー (production環境) </a></div>',
    '<h3>' + name + 'のプレビュー(スポット名表示あり)</h3>',
    '<div><a target="_blank" href="' + previewUrl + showSpotNameParam + '">プレビュー (preview環境) </a></div>',    
    '<div><a target="_blank" href="' + stagingUrl + showSpotNameParam + '">プレビュー (staging環境) </a></div>',
    '<div><a target="_blank" href="' + productionUrl + showSpotNameParam + '">プレビュー (production環境) </a></div>',
    '</div>'
  ]
  var html = HtmlService.createHtmlOutput(content.join(""));
  SpreadsheetApp.getUi().showModelessDialog(html,"preview");
}

function setupS3(){
  var dp     = PropertiesService.getDocumentProperties();
  var key    = dp.getProperty("AWS_KEY");
  var secret = dp.getProperty("AWS_SECRET");
  // Logger.log(key);
  // Logger.log(secret);
  var s3 = S3.getInstance(key, secret);
  return s3;
}

function uploadS3(jsonPath, json) {
  var s3 = setupS3();
  s3.putObject("walk-through-preview", jsonPath, json, {acl: "public-read"});
  Logger.log("Success upload to " + jsonPath);
}

function isWarp(building, spotName) {
  var isWarp = false;
  for (var i=0; i < building.warp_spots.length; i++) {
    if (building.warp_spots[i].name === spotName) {
      isWarp = true;
    }
  }
  return isWarp;
}

// --------------------------------------------------------------------------
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menus = [
    {name: 'AWSアクセスキー設定', functionName: 'setAuth'},
    {name: 'プレビュー表示', functionName: 'preview'},
  ];
  ss.addMenu('ウォークスルーVR', menus);
}

function setAuth(){
  var dp     = PropertiesService.getDocumentProperties();
  var key    = Browser.inputBox('Enter AWS ACCESS KEY', Browser.Buttons.OK_CANCEL);
  var secret = Browser.inputBox('Enter AWS ACCESS SECRET', Browser.Buttons.OK_CANCEL);
  Logger.log(key);
  Logger.log(secret);
  if (key.length > 10 && key != 'cancel') {
    Logger.log('Update AWS_KEY');
    dp.setProperty("AWS_KEY", key);
  }
  if (secret.length > 10 && secret != 'cancel') {
    Logger.log('Update AWS_SECRET');
    dp.setProperty("AWS_SECRET", secret);
  }
}

function preview(){
  Logger.log('Preview publish start');
  var building = generateJson();
  var date = new Date();
  var jsonPath = "building_preview/json/preview_" + date.getTime() + ".json";
  uploadS3(jsonPath, building);
  var jsonUrl = "https://s3-ap-northeast-1.amazonaws.com/walk-through-preview/" + jsonPath;
  showPreviewMsg(building.name, jsonUrl);
  Logger.log('Preview publish end');
}
