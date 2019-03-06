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


        $testname = ilUtil::getASCIIFilename(preg_replace("/\s/", "_", $testname));

        require_once 'Modules/TestQuestionPool/classes/class.ilAssExcelFormatHelper.php';
        $worksheet = new ilAssExcelFormatHelper();
        $worksheet->addSheet($testname);

        $row = 0;
        $col = 0;

		$worksheet->setCell($row, $col++, $testname);

		$row++;
		$col = 0;
		$worksheet->setCell($row, $col++, "startHISsheet");
		$worksheet->setCell($row, $col++, "");
		$worksheet->setCell($row, $col++, "");
		$worksheet->setCell($row, $col++, "endHISsheet");

		$row++;
		$col = 0;
        $worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row);
        $worksheet->setColors('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row, EXCEL_BACKGROUND_COLOR);
        $worksheet->setCell($row, $col++, "mtknr");
        $worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row);
        $worksheet->setColors('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row, EXCEL_BACKGROUND_COLOR);
        $worksheet->setCell($row, $col++, "nachname");
        $worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row);
        $worksheet->setColors('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row, EXCEL_BACKGROUND_COLOR);
        $worksheet->setCell($row, $col++, "vorname");
        $worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row);
        $worksheet->setColors('A' . $row . ':' . $worksheet->getColumnCoord($col) . $row, EXCEL_BACKGROUND_COLOR);
        $worksheet->setCell($row, $col++, "bewertung");


		$data =& $this->getTest()->getCompleteEvaluationData(TRUE);
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
				array_push($datarow2, trim(explode(",", $data->getParticipant($active_id)->getName())[0]));
				array_push($datarow2, trim(explode(",", $data->getParticipant($active_id)->getName())[1]));
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
				$worksheet->setCell($row, $col++, $userfields['matriculation']);
				$worksheet->setCell($row, $col++, trim(explode(",", $data->getParticipant($active_id)->getName())[0]));
				$worksheet->setCell($row, $col++,trim(explode(",", $data->getParticipant($active_id)->getName())[1]));
				$worksheet->setCell($row, $col++, $mark);
			}
		}


		$row++;
		$col = 0;
		$worksheet->setCell($row, $col++,"endHISsheet");
        $excelfile = ilUtil::ilTempnam();
        $worksheet->writeToFile($excelfile);
        mkdir($this->getTest()->getExportDirectory(), 0777, true);
        @copy($excelfile . '.xlsx', $filename->getPathname('xlsx', $testname));
        @unlink($excelfile . '.xlsx');

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

		file_put_contents($filename->getPathname('csv', $testname), $csv);
	}
}