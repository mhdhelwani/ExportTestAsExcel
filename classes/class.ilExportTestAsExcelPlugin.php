<?php
/* Copyright (c) 1998-2013 ILIAS open source, Extended GPL, see docs/LICENSE */

require_once 'Modules/Test/classes/class.ilTestExportPlugin.php';

class ilExportTestAsExcelPlugin extends ilTestExportPlugin
{
    /**
     * Get Plugin Name. Must be same as in class name il<Name>Plugin
     * and must correspond to plugins subdirectory name.
     * Must be overwritten in plugin class of plugin
     * (and should be made final)
     * @return    string    Plugin Name
     */
    function getPluginName()
	{
		return 'ExportTestAsExcel';
	}

	/**
	 * @return string
	 */
	protected function getFormatIdentifier()
	{
		return 'testfile.csv';
	}

	/**
	 * @return string
	 */
	public function getFormatLabel()
	{
		return $this->txt('exporttestasexcel_format');
	}

	protected function buildExportFile(ilTestExportFilename $filename)
	{
		include_once "Services/Excel/classes/class.ilExcelUtils.php";
		$testname = $this->getTest()->getTitle();

		$testSchema = $this->getTest()->mark_schema;

		$rows = array();
		$datarow = array();

		array_push($datarow, $testname);
		array_push($datarow, "");
		array_push($datarow, "");
		array_push($datarow, "");

		array_push($rows, $datarow);

		$datarow = array();
		array_push($datarow, "startHISsheet");
		array_push($datarow, "");
		array_push($datarow, "");
		array_push($datarow, "endHISsheet");

		array_push($rows, $datarow);

		$datarow = array();
		array_push($datarow, "mtknr");
		array_push($datarow, "nachname");
		array_push($datarow, "vorname");
		array_push($datarow, "bewertung");

		array_push($rows, $datarow);


		require_once 'Services/Excel/classes/class.ilExcelWriterAdapter.php';
		$excelfile = ilUtil::ilTempnam();
		$adapter = new ilExcelWriterAdapter($excelfile, FALSE);

		$testname = ilUtil::getASCIIFilename(preg_replace("/\s/", "_", $testname));
		$workbook = $adapter->getWorkbook();
		$workbook->setVersion(8); // Use Excel97/2000 Format
		// Creating a worksheet
		$format_title =& $workbook->addFormat();
		$format_title->setBold();
		$format_title->setColor('black');
		$format_title->setPattern(1);
		$format_title->setFgColor('silver');
		require_once './Services/Excel/classes/class.ilExcelUtils.php';
		$worksheet =& $workbook->addWorksheet(ilExcelUtils::_convert_text($testname));
		$row = 0;
		$col = 0;

		$worksheet->write($row, $col++, ilExcelUtils::_convert_text($testname));

		$row++;
		$col = 0;
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("startHISsheet"));
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text(""));
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text(""));
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("endHISsheet"));

		$row++;
		$col = 0;
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("mtknr"), $format_title);
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("nachname"), $format_title);
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("vorname"), $format_title);
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("bewertung"), $format_title);


		$data =& $this->getTest()->getCompleteEvaluationData(TRUE, $filterby, $filtertext);
		// To check if the full mark is passed or 4
		$mark = $this->getTest()->mark_schema->getMatchingMark(100);
		$mark_short_name = "";
		if (is_object($mark))
		{
			$mark_short_name = $mark->getShortName();
		}

		foreach ($data->getParticipants() as $active_id => $userdata)
		{
			$remove = FALSE;
			if ($passedonly)
			{
				if ($data->getParticipant($active_id)->getPassed() == FALSE)
				{
					$remove = TRUE;
				}
			}
			if (!$remove)
			{
				$datarow2 = array();

				$userfields = ilObjUser::_lookupFields($userdata->getUserID());
				array_push($datarow2, $userfields['matriculation']);
				array_push($datarow2, trim(split(",", $data->getParticipant($active_id)->getName())[0]));
				array_push($datarow2, trim(split(",", $data->getParticipant($active_id)->getName())[1]));
				$mark = $data->getParticipant($active_id)->getMark();
				if ($mark_short_name === "passed" || $mark_short_name === "bestanden"){
					if ($mark === "passed" || $mark === "bestanden")
					{
						$mark = "++";
					}
					else {
						$mark = "--";
					}
				}
				else {
					$mark = str_replace(",", ".", $mark);
					if (is_numeric($mark))
					{
						$mark = $mark * 100;
					}
				}

				array_push($datarow2, $mark);
				array_push($rows, $datarow2);
				$datarow2 = array();

				$row++;
				$col = 0;

				$userfields = ilObjUser::_lookupFields($userdata->getUserID());
				$worksheet->write($row, $col++, ilExcelUtils::_convert_text($userfields['matriculation']));
				$worksheet->write($row, $col++, ilExcelUtils::_convert_text(trim(split(",", $data->getParticipant($active_id)->getName())[0])));
				$worksheet->write($row, $col++, ilExcelUtils::_convert_text(trim(split(",", $data->getParticipant($active_id)->getName())[1])));
				$worksheet->write($row, $col++, ilExcelUtils::_convert_text($mark));
				$counter++;
			}
		}


		$row++;
		$col = 0;
		$worksheet->write($row, $col++, ilExcelUtils::_convert_text("endHISsheet"));
		$workbook->close();

		$datarow = array();
		array_push($datarow, "endHISsheet");
		array_push($datarow, "");
		array_push($datarow, "");
		array_push($datarow, "");

		array_push($rows, $datarow);

		$csv = "";
		$separator = ";";
		foreach ($rows as $evalrow)
		{
			$csvrow =& $this->getTest()->processCSVRow($evalrow, TRUE, $separator);
			$csv .= join($csvrow, $separator) . "\n";
		}

		ilUtil::makeDirParents(dirname($filename->getPathname('csv', $testname)));
		file_put_contents($filename->getPathname('csv', $testname), $csv);

		ilUtil::makeDirParents(dirname($filename->getPathname('xls', $testname)));
		@copy($excelfile, $filename->getPathname('xls', $testname));
		@unlink($excelfile);
	}
}