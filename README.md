EExcelView
----------------------

This is a wrapper of [PHPSpreadsheet](https://phpspreadsheet.readthedocs.io/en/latest/ "PHPSpreadsheet") that extends from CGridView. 
The main idea is to easily export already defined grids to excel files. 
You can use the same array of parameters that the grid uses plus an aditional specific params, which are:

## Usage

The usage is like using CGridView:

```php
$this->widget('EExcelView', array(
     'dataProvider'=> $dataprovider,
     'title' => 'Title',
     'autoWidth' => false,
     'template' => "{summary}\n{items}\n{exportbuttons}\n{pager}",
      ... other options ...
));
```

If you didn't put the extension in the import array (Installation Step 4 below) 
then you need to add the following above the widget call:

~~~
Yii::import('ext.EExcelView');
~~~

### Options you can pass

Document properties

-  creator - The creator of the document
-  title - Title of the excel document
-  subject, description, category

Other params

-  exportType - The file type to export to - defaults to xls.  These are PHPSpreadsheet writer types: Xls, Xlsx, Ods, Csv, Html, Tcpdf, Dompdf, Mpdf .  Can also be passed as a GET var
-  autoWidth - whether to auto extend cells for the content (default true)
-  disablePaging - if set to true, it will export all data (default true)
-  filename - the full path of the filename to export to. If null it will stream to the browser
-  stream - true/false stream to browser or save to file.
-  grid_mode - Whether to display grid or to export it to selected format. Possible values(grid, export)
-  grid_mode_var - defaults to 'grid_mode' GET var for the grid mode
-  sheetTitle - the title of the current sheet.  Default Sheet1
-  rtl_sheet - whether to set the sheet as right-to-left

## Requirements

Yii 1.1, Composer.

## Installation

### Step 1 - Create/update your composer.json file

Create or update your composer.json inside the protected/ folder. 
Inside there, add the following line to the require section of composer.json:

````
  "require": {
    ...
    "phpoffice/phpspreadsheet": ">= 1.10.1"
    ...
  },
````

Then run composer update.

### Step 2 - Add composer's autoloader (if you've not already done so)

Inside your entrypoint (index.php in the root of the project), just above where the Yii framework is require_once'd, add:
````
// Autoload including composer
require_once(__DIR__ . '/protected/vendor/autoload.php');
````

### Step 3 - Download and install this extension
 
Download this extension and put it in extensions/EExcelView folder.

### Step 4. (Optional) Add the path to your config file

Add the following to the import section in your config file:

~~~
[php]

		// autoloading model and component classes
		'import'=>array(
			.............
			'ext.EExcelView.*',
		),
~~~

## Release notes

### ver 0.4
- Migrated to use PHPSpreadsheet rather than PHPExcel
- Migrated to use composer

#### ver 0.32
- Fixed the CButtonColumn bug "Property "CButtonColumn"."name" is not defined"...
- Fixed the issues with multiple level of buffering(le top)
- Fixed export buttons column bug(le top)
- Added documetation of all properties in the extension page.

#### ver 0.31
Mainly bug fix release. I also modified the extension to be able to run in console mode(although its a widget). I realy needed something to send me reports in xls. Here is the code snippet:


~~~
[php]
  	$factory = new CWidgetFactory();	
		$widget = $factory->createWidget($this, 'EExcelView', array(
			'dataProvider'=>$provider,
			'grid_mode'=>'export',
			'title'=>'Title',
			'filename'=>'report.xlsx',
			'stream'=>false,
			'exportType'=>'Excel2007',
			'columns'=>array(
				'col1',
				'col2',
			),
		));
		
		$widget->init();
		$widget->run();
~~~


#### ver 0.3
Major change EExceView now can fully replace CGridView. There is an optional parameter 'gridmode'(default "grid") and when not set it acts just like normal grid. If gridmode="export" it exports the grid in the selected "exportType"(default "Excel5"). The path to PHPExcel can now be customized via "libPath" variable. If the lib is not found a warning is generated and export is disabled.
 
- Added: "gridmode" possible values(grid, export). When set to grid acts like normal grid.
- Added: "libPath" property, defaults to "application.extensions.phpexcel.Classes.PHPExcel". The location of PHP Excel lib.
- Added: "grid_mode_var" the GET variable name which can be used to change the grid mode. defaults to "grid_mode"
- Added: new template item "exportbuttons" which renders export links. Example: 
....
'template'=>"{summary}\n{items}\n{exportbuttons}\n{pager}"
....
- Added: config for export buttons.
  "exportButtonsCSS"(default="summary") css class for the div which wraps the buttons. "exportButtons" - array of exportTypes, for each item it renders link which exports the grid. If you specify an array you can customize the caption of the link. All export type have default captions. Example: array('Excel5'=>array('caption'=>'*.xls'), PDF). Property exportText which defines the text before the links, deault = "Export to: ".
- Added: renders CLinkColumns. Just the text of the links.
- Added: support for footer, just like normal footer.
- Added: callback functions called after processing of the cell.$cell is the php excel cell object. Refer to php excel documentation for help. $data is the same var the grids uses for expressions. $value is the text the extension had put in the cell : 
1. onRenderHeaderCell(PHPExcel_Cell $cell, string $value)
2. onRenderDataCell(PHPExcel_Cell $cell, $data, string $value)
3. onRenderFooterCell(PHPExcel_Cell $cell, string $value)

Thanks to all who have written me with ideas how to extend this extension. Sorry by I've been realy busy lately so I didnt have time to bring this earlier. I have tested the extension but if you find any bug or something you'd like me to add, feel free to write. Happy yiister :) 

#### ver 0.2
- Fixed: Error when saving to filename.
- Fixed: bug in column indexing when > 26
- Added: support for CSV
- Added: new param "stream"(default true) when set to true streams to browser else save to local filename
- Added: get real column headers
- Added: get title from page title when not specified
