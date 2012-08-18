EExcelView
----------------------

[Forum thread](http://www.yiiframework.com/forum/index.php/topic/18119-extensioneexcelview "Forum thread")

This is my first extension ever :). It is a wrapper of [PHPExcel](http://phpexcel.codeplex.com/ "PHPExcel") that extends from CGrivView. The main idea is to easily export already defined grids to excel files. You can use the same array of parameters that the grid uses plus an aditional specific params, which are:

##Configuraiton

Document properties

-  creator - The creator of the document
-  title - Title of the excel document
-  subject, description, category

Other params

-  libPath - path to PHPExcel, defaults to 'ext.phpexcel.Classes.PHPExcel'
-  autoWidth - whether to auto extend cells for the content(default true)
-  exportType - the type of the export, all possible types of PHPExcel lib(Excel5, Excel2007,PDF, HTML)
-  disablePaging - if set to true, it will export all data (default true)
-  filename - the full path of the filename to export to. If null it will stream to the browser
-  stream - true/false stream to browser or save to file.
-  grid_mode - Whether to display grid or to export it to selected format. Possible values(grid, export)
-  grid_mode_var - defaults to 'grid_mode' GET var for the grid mode

##Requirements

Yii 1.1 or above and [PHPExcel](http://phpexcel.codeplex.com/ "PHPExcel") library

##Usage

The usage is like using CGridView:
~~~
[php]
$this->widget('EExcelView', array(
     'dataProvider'=> $dataprovider,
     'title'=>'Title',
     'autoWidth'=>false,
      ..... other options 
));
~~~

##Note
The path of the PHPExcel library is hardcoded to: **application.extensions.phpexcel.Classes.PHPExcel**

Asuming that you have extracted the lib under extensions folder. I will change that later.

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