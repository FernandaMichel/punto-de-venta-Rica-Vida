<?php


namespace PhpOffice\PhpWord\Reader;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Drawing;
use PhpOffice\PhpWord\Shared\OLERead;
use PhpOffice\PhpWord\Style;


class MsDoc extends AbstractReader implements ReaderInterface
{
   
    private $phpWord;

   
    private $dataWorkDocument;
    
    private $data1Table;
   
    private $dataData;
    
    private $dataObjectPool;
    
    private $arrayCharacters = array();
    
    private $arrayFib = array();
    
    private $arrayFonts = array();
    
    private $arrayParagraphs = array();
    
    private $arraySections = array();

    const VERSION_97 = '97';
    const VERSION_2000 = '2000';
    const VERSION_2002 = '2002';
    const VERSION_2003 = '2003';
    const VERSION_2007 = '2007';

    const SPRA_VALUE = 10;
    const SPRA_VALUE_OPPOSITE = 20;

    const OFFICEARTBLIPEMF = 0xF01A;
    const OFFICEARTBLIPWMF = 0xF01B;
    const OFFICEARTBLIPPICT = 0xF01C;
    const OFFICEARTBLIPJPG = 0xF01D;
    const OFFICEARTBLIPPNG = 0xF01E;
    const OFFICEARTBLIPDIB = 0xF01F;
    const OFFICEARTBLIPTIFF = 0xF029;
    const OFFICEARTBLIPJPEG = 0xF02A;

    const MSOBLIPERROR = 0x00;
    const MSOBLIPUNKNOWN = 0x01;
    const MSOBLIPEMF = 0x02;
    const MSOBLIPWMF = 0x03;
    const MSOBLIPPICT = 0x04;
    const MSOBLIPJPEG = 0x05;
    const MSOBLIPPNG = 0x06;
    const MSOBLIPDIB = 0x07;
    const MSOBLIPTIFF = 0x11;
    const MSOBLIPCMYKJPEG = 0x12;

   
    public function load($filename)
    {
        $this->phpWord = new PhpWord();

        $this->loadOLE($filename);

        $this->readFib($this->dataWorkDocument);
        $this->readFibContent();

        return $this->phpWord;
    }

    
    private function loadOLE($filename)
    {
        $ole = new OLERead();
        $ole->read($filename);

        $this->dataWorkDocument = $ole->getStream($ole->wrkdocument);
        $this->data1Table = $ole->getStream($ole->wrk1Table);
        $this->dataData = $ole->getStream($ole->wrkData);
        $this->dataObjectPool = $ole->getStream($ole->wrkObjectPool);
        $this->_SummaryInformation = $ole->getStream($ole->summaryInformation);
        $this->_DocumentSummaryInformation = $ole->getStream($ole->documentSummaryInformation);
    }

    private function getNumInLcb($lcb, $iSize)
    {
        return ($lcb - 4) / (4 + $iSize);
    }

    private function getArrayCP($data, $posMem, $iNum)
    {
        $arrayCP = array();
        for ($inc = 0; $inc < $iNum; $inc++) {
            $arrayCP[$inc] = self::getInt4d($data, $posMem);
            $posMem += 4;
        }
        return $arrayCP;
    }

   
    private function readFib($data)
    {
        $pos = 0;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 4;
        $pos += 1;
        $pos += 1;
        $pos += 2;
        $pos += 2;
        $pos += 4;
        $pos += 4;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 2;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $this->arrayFib['ccpText'] = self::getInt4d($data, $pos);
        $pos += 4;
        $this->arrayFib['ccpFtn'] = self::getInt4d($data, $pos);
        $pos += 4;
        $this->arrayFib['ccpHdd'] = self::getInt4d($data, $pos);
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;
        $pos += 4;

        $cbRgFcLcb = self::getInt2d($data, $pos);
        $pos += 2;
        switch ($cbRgFcLcb) {
            case 0x005D:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                break;
            case 0x006C:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                break;
            case 0x0088:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2002);
                break;
            case 0x00A4:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2002);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2003);
                break;
            case 0x00B7:
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_97);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2000);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2002);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2003);
                $pos = $this->readBlockFibRgFcLcb($data, $pos, self::VERSION_2007);
                break;
        }
        $this->arrayFib['cswNew'] = self::getInt2d($data, $pos);
        $pos += 2;

        if ($this->arrayFib['cswNew'] != 0) {
        }

        return $pos;
    }

    private function readBlockFibRgFcLcb($data, $pos, $version)
    {
        if ($version == self::VERSION_97) {
            $this->arrayFib['fcStshfOrig'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbStshfOrig'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcStshf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbStshf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffndRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffndRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffndTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffndTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfandRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfandRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfandTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfandTxt '] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfSed'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfSed'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcPad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcPad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfPhe'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfPhe'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfGlsy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfHdd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfHdd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBteChpx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBteChpx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBtePapx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBtePapx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfSea'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfSea'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfFfn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfFfn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldAtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldAtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcCmds'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbCmds'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfMcr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPrDrvr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPrDrvr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPrEnvPort'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPrEnvPort'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPrEnvLand'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPrEnvLand'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcWss'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbWss'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDop'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDop'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfAssoc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfAssoc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcClx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbClx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAutosaveSource'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAutosaveSource'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcGrpXstAtnOwners'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbGrpXstAtnOwners'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfAtnBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfAtnBkmk'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcSpaMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcSpaMom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcSpaHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcSpaHdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfAtnBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfAtnBkf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfAtnBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfAtnBkl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPms'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPms'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcFormFldSttbs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbFormFldSttbs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfendRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfendRef'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfendTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfendTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDggInfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDggInfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfRMark'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfRMark'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfAutoCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfAutoCaption'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfWkb'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfWkb'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfSpl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfSpl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcftxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcftxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfFldTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfFldTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfHdrtxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfHdrtxbxTxt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffldHdrTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffldHdrTxbx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcStwUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbStwUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbTtmbd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbTtmbd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcCookieData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbCookieData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdMotherOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdFtnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdEdnOldOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfIntlFld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfIntlFld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRouteSlip'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRouteSlip'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbSavedBy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbSavedBy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbFnm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbFnm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlfLst'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlfLst'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlfLfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlfLfo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfTxbxBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfTxbxBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfTxbxHdrBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfTxbxHdrBkd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDocUndoWord9'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDocUndoWord9'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUskf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUskf'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcupcRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcupcRgbUse'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcupcUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcupcUsp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbGlsyStyle'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbGlsyStyle'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlgosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlgosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcocx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcocx'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBteLvc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBteLvc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['dwLowDateTime'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['dwHighDateTime'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfLvcPre10'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfLvcPre10'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfAsumy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfAsumy'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfGram'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfGram'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbListNames'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbListNames'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfUssr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfUssr'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2000) {
            $this->arrayFib['fcPlcfTch'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfTch'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRmdThreading'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRmdThreading'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcMid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbMid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbRgtplc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbRgtplc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcMsoEnvelope'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbMsoEnvelope'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfLad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfLad'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcRgDofr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbRgDofr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcosl'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfCookieOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfCookieOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdMotherOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdFtnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdEdnOld'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2002) {
            $this->arrayFib['fcUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfPgp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfPgp'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfuim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfuim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlfguidUim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlfguidUim'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAtrdExtra'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAtrdExtra'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlrsid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlrsid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfcookie'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfcookie'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklFactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcFactoidData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbFactoidData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcDocUndo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbDocUndo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklFcc'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfbkmkBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfbkmkBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfbkfBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfbkfBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfbklBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfbklBPRepairs'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPmsNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPmsNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcODSO'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbODSO'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcffactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcffactoid'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcOldXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcNewXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcMixedXP'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2003) {
            $this->arrayFib['fcHplxsdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbHplxsdr'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklSdt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcCustomXForm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbCustomXForm'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklProt'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbProtUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbProtUser'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfpmiNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfpmiNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcOld'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcOldInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcNew'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcflvcNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcflvcNewInline'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfdMother'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfdFtn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPgdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPgdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcBkdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbBkdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfdEdn'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcAfd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbAfd'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        if ($version == self::VERSION_2007) {
            $this->arrayFib['fcPlcfmthd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfmthd'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklMoveFrom'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklMoveTo'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused1'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused2'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused3'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcSttbfBkmkArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbSttbfBkmkArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBkfArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBkfArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcPlcfBklArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbPlcfBklArto'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcArtoData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbArtoData'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused4'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused5'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused5'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcUnused6'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbUnused6'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcOssTheme'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbOssTheme'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['fcColorSchemeMapping'] = self::getInt4d($data, $pos);
            $pos += 4;
            $this->arrayFib['lcbColorSchemeMapping'] = self::getInt4d($data, $pos);
            $pos += 4;
        }
        return $pos;
    }

    private function readFibContent()
    {
        $this->readRecordSttbfFfn();

        $this->readRecordPlcfSed();

       $this->readRecordPlcfBtePapx();

                $this->readRecordPlcfBteChpx();

        $this->generatePhpWord();
    }


    private function readRecordPlcfSed()
    {
        $posMem = $this->arrayFib['fcPlcfSed'];
        
        $aCP = array();
        $aCP[0] = self::getInt4d($this->data1Table, $posMem);
        $posMem += 4;
        $aCP[1] = self::getInt4d($this->data1Table, $posMem);
        $posMem += 4;

        
        $numSed = $this->getNumInLcb($this->arrayFib['lcbPlcfSed'], 12);

        $aSed = array();
        for ($iInc = 0; $iInc < $numSed; ++$iInc) {
            
            $posMem += 2;
            $aSed[$iInc] = self::getInt4d($this->data1Table, $posMem);
            $posMem += 4;
            $posMem += 2;
            $posMem += 4;
        }

        foreach ($aSed as $offsetSed) {
            $cb = self::getInt2d($this->dataWorkDocument, $offsetSed);
            $offsetSed += 2;

            $oStylePrl = $this->readPrl($this->dataWorkDocument, $offsetSed, $cb);
            $offsetSed += $oStylePrl->length;

            $this->arraySections[] = $oStylePrl;
        }
    }

    
    private function readRecordSttbfFfn()
    {
        $posMem = $this->arrayFib['fcSttbfFfn'];

        $cData = self::getInt2d($this->data1Table, $posMem);
        $posMem += 2;
        $cbExtra = self::getInt2d($this->data1Table, $posMem);
        $posMem += 2;

        if ($cData < 0x7FF0 && $cbExtra == 0) {
            for ($inc = 0; $inc < $cData; $inc++) {
                $posMem += 1;
                $posMem += 1;
                $posMem += 2;
                $posMem += 1;
                $ixchSzAlt = self::getInt1d($this->data1Table, $posMem);
                $posMem += 1;
                $posMem += 10;
                $posMem += 24;
                $xszFfn = '';
                do {
                    $char = self::getInt2d($this->data1Table, $posMem);
                    $posMem += 2;
                    if ($char > 0) {
                        $xszFfn .= chr($char);
                    }
                } while ($char != 0);
                $xszAlt = '';
                if ($ixchSzAlt > 0) {
                    do {
                        $char = self::getInt2d($this->data1Table, $posMem);
                        $posMem += 2;
                        if ($char == 0) {
                            break;
                        }
                        $xszAlt .= chr($char);
                    } while ($char != 0);
                }
                $this->arrayFonts[] = array(
                    'main' => $xszFfn,
                    'alt' => $xszAlt,
                );
            }
        }
    }

    
    private function readRecordPlcfBtePapx()
    {
        $posMem = $this->arrayFib['fcPlcfBtePapx'];
        $num = $this->getNumInLcb($this->arrayFib['lcbPlcfBtePapx'], 4);
        $posMem += 4 * ($num + 1);
        $arrAPnBtePapx = $this->getArrayCP($this->data1Table, $posMem, $num);
        $posMem += 4 * $num;

        foreach ($arrAPnBtePapx as $aPnBtePapx) {
            $offsetBase = $aPnBtePapx * 512;
            $offset = $offsetBase;

            $string = '';

            $numRun = self::getInt1d($this->dataWorkDocument, $offset + 511);
            $arrayRGFC = array();
            for ($inc = 0; $inc <= $numRun; $inc++) {
                $arrayRGFC[$inc] = self::getInt4d($this->dataWorkDocument, $offset);
                $offset += 4;
            }
            $arrayRGB = array();
            for ($inc = 1; $inc <= $numRun; $inc++) {
                $arrayRGB[$inc] = self::getInt1d($this->dataWorkDocument, $offset);
                $offset += 1;
                $offset += 12;
            }

            foreach (array_keys($arrayRGFC) as $key) {
                if (!isset($arrayRGFC[$key + 1])) {
                    break;
                }
                $strLen = $arrayRGFC[$key + 1] - $arrayRGFC[$key] - 1;
                for ($inc = 0; $inc < $strLen; $inc++) {
                    $byte = self::getInt1d($this->dataWorkDocument, $arrayRGFC[$key] + $inc);
                    if ($byte > 0) {
                        $string .= chr($byte);
                    }
                }
            }
            $this->arrayParagraphs[] = $string;

        }
    }

   
    private function readRecordPlcfBteChpx()
    {
        $posMem = $this->arrayFib['fcPlcfBteChpx'];
        $num = $this->getNumInLcb($this->arrayFib['lcbPlcfBteChpx'], 4);
        $aPnBteChpx = array();
        for ($inc = 0; $inc <= $num; $inc++) {
            $aPnBteChpx[$inc] = self::getInt4d($this->data1Table, $posMem);
            $posMem += 4;
        }
        $pnFkpChpx = self::getInt4d($this->data1Table, $posMem);
        $posMem += 4;

        $offsetBase = $pnFkpChpx * 512;
        $offset = $offsetBase;

        $numRGFC = self::getInt1d($this->dataWorkDocument, $offset + 511);
        $arrayRGFC = array();
        for ($inc = 0; $inc <= $numRGFC; $inc++) {
            $arrayRGFC[$inc] = self::getInt4d($this->dataWorkDocument, $offset);
            $offset += 4;
        }

        $arrayRGB = array();
        for ($inc = 1; $inc <= $numRGFC; $inc++) {
            $arrayRGB[$inc] = self::getInt1d($this->dataWorkDocument, $offset);
            $offset += 1;
        }

        $start = 0;
        foreach ($arrayRGB as $keyRGB => $rgb) {
            $oStyle = new \stdClass();
            $oStyle->pos_start = $start;
            $oStyle->pos_len = (int)ceil((($arrayRGFC[$keyRGB] -1) - $arrayRGFC[$keyRGB -1]) / 2);
            $start += $oStyle->pos_len;

            if ($rgb > 0) {
                               $posRGB = $offsetBase + $rgb * 2;

                $cb = self::getInt1d($this->dataWorkDocument, $posRGB);
                $posRGB += 1;

                $oStyle->style = $this->readPrl($this->dataWorkDocument, $posRGB, $cb);
                $posRGB += $oStyle->style->length;
            }
            $this->arrayCharacters[] = $oStyle;
        }
    }

   
    private function readSprm($sprm)
    {
        $oSprm = new \stdClass();
        $oSprm->isPmd = $sprm & 0x01FF;
        $oSprm->f = ($sprm / 512) & 0x0001;
        $oSprm->sgc = ($sprm / 1024) & 0x0007;
        $oSprm->spra = ($sprm / 8192);
        return $oSprm;
    }

   
    private function readSprmSpra($data, $pos, $oSprm)
    {
        $length = 0;
        $operand = null;

        switch(dechex($oSprm->spra)) {
            case 0x0:
                $operand = self::getInt1d($data, $pos);
                $length = 1;
                switch(dechex($operand)) {
                    case 0x00:
                        $operand = false;
                        break;
                    case 0x01:
                        $operand = true;
                        break;
                    case 0x80:
                        $operand = self::SPRA_VALUE;
                        break;
                    case 0x81:
                        $operand = self::SPRA_VALUE_OPPOSITE;
                        break;
                }
                break;
            case 0x1:
                $operand = self::getInt1d($data, $pos);
                $length = 1;
                break;
            case 0x2:
            case 0x4:
            case 0x5:
                $operand = self::getInt2d($data, $pos);
                $length = 2;
                break;
            case 0x3:
                if ($oSprm->isPmd != 0x70) {
                    $operand = self::getInt4d($data, $pos);
                    $length = 4;
                }
                break;
            case 0x7:
                $operand = self::getInt3d($data, $pos);
                $length = 3;
                break;
            default:
        }

        return array(
            'length' => $length,
            'operand' => $operand,
        );
    }

    private function readPrl($data, $pos, $cbNum)
    {
        $posStart = $pos;
        $oStylePrl = new \stdClass();

        $sprmCPicLocation = null;
        $sprmCFData = null;
        $sprmCFSpec = null;

        do {
            $operand = null;

            $sprm = self::getInt2d($data, $pos);
            $oSprm = $this->readSprm($sprm);
            $pos += 2;
            $cbNum -= 2;

            $arrayReturn = $this->readSprmSpra($data, $pos, $oSprm);
            $pos += $arrayReturn['length'];
            $cbNum -= $arrayReturn['length'];
            $operand = $arrayReturn['operand'];

            switch(dechex($oSprm->sgc)) {
                case 0x01:
                    break;
                case 0x02:
                    if (!isset($oStylePrl->styleFont)) {
                        $oStylePrl->styleFont = array();
                    }
                    switch($oSprm->isPmd) {
                        case 0x01:
                            break;
                        case 0x02:
                            break;
                        case 0x03:
                            $sprmCPicLocation = $operand;
                            break;
                        case 0x06:
                            $sprmCFData = dechex($operand) == 0x00 ? false : true;
                            break;
                        case 0x36:
                            switch($operand) {
                                case false:
                                case true:
                                    $oStylePrl->styleFont['italic'] = $operand;
                                    break;
                                case self::SPRA_VALUE:
                                    $oStylePrl->styleFont['italic'] = false;
                                    break;
                                case self::SPRA_VALUE_OPPOSITE:
                                    $oStylePrl->styleFont['italic'] = true;
                                    break;
                            }
                            break;
                        case 0x30:
                            break;
                        case 0x35:
                            switch($operand) {
                                case false:
                                case true:
                                    $oStylePrl->styleFont['bold'] = $operand;
                                    break;
                                case self::SPRA_VALUE:
                                    $oStylePrl->styleFont['bold'] = false;
                                    break;
                                case self::SPRA_VALUE_OPPOSITE:
                                    $oStylePrl->styleFont['bold'] = true;
                                    break;
                            }
                            break;
                        case 0x37:
                            switch($operand) {
                                case false:
                                case true:
                                    $oStylePrl->styleFont['strikethrough'] = $operand;
                                    break;
                                case self::SPRA_VALUE:
                                    $oStylePrl->styleFont['strikethrough'] = false;
                                    break;
                                case self::SPRA_VALUE_OPPOSITE:
                                    $oStylePrl->styleFont['strikethrough'] = true;
                                    break;
                            }
                            break;
                        case 0x3E:
                            switch(dechex($operand)) {
                                case 0x00:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_NONE;
                                    break;
                                case 0x01:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_SINGLE;
                                    break;
                                case 0x02:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WORDS;
                                    break;
                                case 0x03:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOUBLE;
                                    break;
                                case 0x04:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTTED;
                                    break;
                                case 0x06:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_HEAVY;
                                    break;
                                case 0x07:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASH;
                                    break;
                                case 0x09:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTHASH;
                                    break;
                                case 0x0A:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTDOTDASH;
                                    break;
                                case 0x0B:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WAVY;
                                    break;
                                case 0x14:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTTEDHEAVY;
                                    break;
                                case 0x17:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASHHEAVY;
                                    break;
                                case 0x19:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTHASHHEAVY;
                                    break;
                                case 0x1A:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DOTDOTDASHHEAVY;
                                    break;
                                case 0x1B:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WAVYHEAVY;
                                    break;
                                case 0x27:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASHLONG;
                                    break;
                                case 0x2B:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_WAVYDOUBLE;
                                    break;
                                case 0x37:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_DASHLONGHEAVY;
                                    break;
                                default:
                                    $oStylePrl->styleFont['underline'] = Style\Font::UNDERLINE_NONE;
                                    break;
                            }
                            break;
                                case 0x42:
                            switch(dechex($operand)) {
                                case 0x00:
                                case 0x01:
                                    $oStylePrl->styleFont['color'] = '000000';
                                    break;
                                case 0x02:
                                    $oStylePrl->styleFont['color'] = '0000FF';
                                    break;
                                case 0x03:
                                    $oStylePrl->styleFont['color'] = '00FFFF';
                                    break;
                                case 0x04:
                                    $oStylePrl->styleFont['color'] = '00FF00';
                                    break;
                                case 0x05:
                                    $oStylePrl->styleFont['color'] = 'FF00FF';
                                    break;
                                case 0x06:
                                    $oStylePrl->styleFont['color'] = 'FF0000';
                                    break;
                                case 0x07:
                                    $oStylePrl->styleFont['color'] = 'FFFF00';
                                    break;
                                case 0x08:
                                    $oStylePrl->styleFont['color'] = 'FFFFFF';
                                    break;
                                case 0x09:
                                    $oStylePrl->styleFont['color'] = '000080';
                                    break;
                                case 0x0A:
                                    $oStylePrl->styleFont['color'] = '008080';
                                    break;
                                case 0x0B:
                                    $oStylePrl->styleFont['color'] = '008000';
                                    break;
                                case 0x0C:
                                    $oStylePrl->styleFont['color'] = '800080';
                                    break;
                                case 0x0D:
                                    $oStylePrl->styleFont['color'] = '800080';
                                    break;
                                case 0x0E:
                                    $oStylePrl->styleFont['color'] = '808000';
                                    break;
                                case 0x0F:
                                    $oStylePrl->styleFont['color'] = '808080';
                                    break;
                                case 0x10:
                                    $oStylePrl->styleFont['color'] = 'C0C0C0';
                            }
                            break;
                        case 0x43:
                            $oStylePrl->styleFont['size'] = dechex($operand/2);
                            break;
                        case 0x48:
                            if (!isset($oStylePrl->styleFont['superScript'])) {
                                $oStylePrl->styleFont['superScript'] = false;
                            }
                            if (!isset($oStylePrl->styleFont['subScript'])) {
                                $oStylePrl->styleFont['subScript'] = false;
                            }
                            switch (dechex($operand)) {
                                case 0x00:
                                    break;
                                case 0x01:
                                    $oStylePrl->styleFont['superScript'] = true;
                                    break;
                                case 0x02:
                                    $oStylePrl->styleFont['subScript'] = true;
                                    break;
                            }
                            break;
                        case 0x4F:
                            $oStylePrl->styleFont['name'] = '';
                            if (isset($this->arrayFonts[$operand])) {
                                $oStylePrl->styleFont['name'] = $this->arrayFonts[$operand]['main'];
                            }
                            break;
                        case 0x50:
                            break;
                        case 0x51:
                            break;
                        case 0x55:
                            $sprmCFSpec = $operand;
                            break;
                        case 0x5E:
                            break;
                        case 0x5D:
                            break;
                        case 0x61:
                            break;
                        case 0x66:
                            $pos += 2;
                            $cbNum -= 2;
                            break;
                        case 0x70:
                            $red = str_pad(dechex(self::getInt1d($this->dataWorkDocument, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $green = str_pad(dechex(self::getInt1d($this->dataWorkDocument, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $blue = str_pad(dechex(self::getInt1d($this->dataWorkDocument, $pos)), 2, '0', STR_PAD_LEFT);
                            $pos += 1;
                            $pos += 1;
                            $oStylePrl->styleFont['color'] = $red.$green.$blue;
                            $cbNum -= 4;
                            break;
                        default:
                            
                    }
                    break;
                case 0x03:
                    break;
                case 0x04:
                    if (!isset($oStylePrl->styleSection)) {
                        $oStylePrl->styleSection = array();
                    }
                    switch($oSprm->isPmd) {
                        case 0x0E:
                            break;
                        case 0x1F:
                            $oStylePrl->styleSection['pageSizeW'] = $operand;
                            break;
                        case 0x20:
                            $oStylePrl->styleSection['pageSizeH'] = $operand;
                            break;
                        case 0x21:
                            $oStylePrl->styleSection['marginLeft'] = $operand;
                            break;
                        case 0x22:
                            $oStylePrl->styleSection['marginRight'] = $operand;
                            break;
                        case 0x23:
                            $oStylePrl->styleSection['marginTop'] = $operand;
                            break;
                        case 0x24:
                            $oStylePrl->styleSection['marginBottom'] = $operand;
                            break;
                        case 0x28:
                            break;
                        case 0x30:
                            break;
                        
                        case 0x31:
                            break;
                        case 0x32:
                            break;
                        case 0x33:
                            break;
                        default:
                            
                    }
                    break;
                case 0x05:
                    break;
            }
        } while ($cbNum > 0);

        if (!is_null($sprmCPicLocation)) {
            if (!is_null($sprmCFData) && $sprmCFData == 0x01) {
                
            } else {
                
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 2;
                $mfpfMm = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 4;
                $picmidDxaGoal = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidDyaGoal = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidMx = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidMy = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidDxaCropLeft = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidDxaCropTop = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidDxaCropRight = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $picmidDxaCropBottom = self::getInt2d($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 1;
                $sprmCPicLocation += 1;
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 4;
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 2;
                $sprmCPicLocation += 2;

                if ($mfpfMm == 0x0066) {
                    $cchPicName = self::getInt1d($this->dataData, $sprmCPicLocation);
                    $sprmCPicLocation += 1;

                    $stPicName = '';
                    for ($inc = 0; $inc <= $cchPicName; $inc++) {
                        $chr = self::getInt1d($this->dataData, $sprmCPicLocation);
                        $sprmCPicLocation += 1;
                        $stPicName .= chr($chr);
                    }
                }

               
                $shapeRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
                $sprmCPicLocation += 8;
                if ($shapeRH['recVer'] == 0xF && $shapeRH['recInstance'] == 0x000 && $shapeRH['recType'] == 0xF004) {
                    $sprmCPicLocation += $shapeRH['recLen'];
                }
                
                $fileBlockRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
                while ($fileBlockRH['recType'] == 0xF007 || ($fileBlockRH['recType'] >= 0xF018 && $fileBlockRH['recType'] <= 0xF117)) {
                    $sprmCPicLocation += 8;
                    switch ($fileBlockRH['recType']) {
                        
                        case 0xF007:
                            $sprmCPicLocation += 1;
                            $sprmCPicLocation += 1;
                            $sprmCPicLocation += 16;
                            $sprmCPicLocation += 2;
                            $sprmCPicLocation += 4;
                            $sprmCPicLocation += 4;
                            $sprmCPicLocation += 4;
                            $sprmCPicLocation += 1;
                            $cbName = self::getInt1d($this->dataData, $sprmCPicLocation);
                            $sprmCPicLocation += 1;
                            $sprmCPicLocation += 1;
                            $sprmCPicLocation += 1;
                            if ($cbName > 0) {
                                $nameData = '';
                                for ($inc = 0; $inc <= ($cbName / 2); $inc++) {
                                    $chr = self::getInt2d($this->dataData, $sprmCPicLocation);
                                    $sprmCPicLocation += 2;
                                    $nameData .= chr($chr);
                                }
                            }
                            
                            $embeddedBlipRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
                            switch ($embeddedBlipRH['recType']) {
                                case self::OFFICEARTBLIPJPG:
                                case self::OFFICEARTBLIPJPEG:
                                    if (!isset($oStylePrl->image)) {
                                        $oStylePrl->image = array();
                                    }
                                    $sprmCPicLocation += 8;
                                    $sprmCPicLocation += 16;
                                    if ($embeddedBlipRH['recInstance'] == 0x6E1) {
                                        $sprmCPicLocation += 16;
                                    }
                                    $sprmCPicLocation += 1;
                                    $oStylePrl->image['data'] = substr($this->dataData, $sprmCPicLocation, $embeddedBlipRH['recLen']);
                                    $oStylePrl->image['format'] = 'jpg';
                                    $iCropWidth = $picmidDxaGoal - ($picmidDxaCropLeft + $picmidDxaCropRight);
                                    $iCropHeight = $picmidDyaGoal - ($picmidDxaCropTop + $picmidDxaCropBottom);
                                    if (!$iCropWidth) {
                                        $iCropWidth = 1;
                                    }
                                    if (!$iCropHeight) {
                                        $iCropHeight = 1;
                                    }
                                    $oStylePrl->image['width'] = Drawing::twipsToPixels($iCropWidth * $picmidMx / 1000);
                                    $oStylePrl->image['height'] = Drawing::twipsToPixels($iCropHeight * $picmidMy / 1000);

                                    $sprmCPicLocation += $embeddedBlipRH['recLen'];
                                    break;
                                default:
                            }
                            break;
                    }
                    $fileBlockRH = $this->loadRecordHeader($this->dataData, $sprmCPicLocation);
                }
            }
        }

        $oStylePrl->length = $pos - $posStart;
        return $oStylePrl;
    }

    
    private function loadRecordHeader($stream, $pos)
    {
        $rec = self::getInt2d($stream, $pos);
        $recType = self::getInt2d($stream, $pos + 2);
        $recLen = self::getInt4d($stream, $pos + 4);
        return array(
            'recVer' => ($rec >> 0) & bindec('1111'),
            'recInstance' => ($rec >> 4) & bindec('111111111111'),
            'recType' => $recType,
            'recLen' => $recLen,
        );
    }

    private function generatePhpWord()
    {
        foreach ($this->arraySections as $itmSection) {
            $oSection = $this->phpWord->addSection();
            $oSection->setSettings($itmSection->styleSection);

            $sHYPERLINK = '';
            foreach ($this->arrayParagraphs as $itmParagraph) {
                $textPara = $itmParagraph;
                foreach ($this->arrayCharacters as $oCharacters) {
                    $subText = substr($textPara, $oCharacters->pos_start, $oCharacters->pos_len);
                    $subText = str_replace(chr(13), PHP_EOL, $subText);
                    $arrayText = explode(PHP_EOL, $subText);
                    if (end($arrayText) == '') {
                        array_pop($arrayText);
                    }
                    if (reset($arrayText) == '') {
                        array_shift($arrayText);
                    }

                    $styleFont = array();
                    if (isset($oCharacters->style)) {
                        if (isset($oCharacters->style->styleFont)) {
                            $styleFont = $oCharacters->style->styleFont;
                        }
                    }

                    foreach ($arrayText as $sText) {
                        if (empty($sText) && !empty($sHYPERLINK)) {
                            $arrHYPERLINK = explode('"', $sHYPERLINK);
                            $oSection->addLink($arrHYPERLINK[1], null);
                            $sHYPERLINK = '';
                        }

                        if (empty($sText)) {
                            $oSection->addTextBreak();
                            $sHYPERLINK = '';
                        }

                        if (!empty($sText)) {
                            if (!empty($sHYPERLINK) && ord($sText[0]) > 20) {
                                $sHYPERLINK .= $sText;
                            }
                            if (empty($sHYPERLINK)) {
                                if (ord($sText[0]) > 20) {
                                    if (strpos(trim($sText), 'HYPERLINK "') === 0) {
                                        $sHYPERLINK = $sText;
                                    } else {
                                        $oSection->addText($sText, $styleFont);
                                    }
                                }
                                if (ord($sText[0]) == 1) {
                                    if (isset($oCharacters->style->image)) {
                                        $fileImage = tempnam(sys_get_temp_dir(), 'PHPWord_MsDoc').'.'.$oCharacters->style->image['format'];
                                        file_put_contents($fileImage, $oCharacters->style->image['data']);
                                        $oSection->addImage($fileImage, array('width' => $oCharacters->style->image['width'], 'height' => $oCharacters->style->image['height']));
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
    }

    
    public static function getInt1d($data, $pos)
    {
        return ord($data[$pos]);
    }

    
    public static function getInt2d($data, $pos)
    {
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8);
    }

    
    public static function getInt3d($data, $pos)
    {
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16);
    }

    
    public static function getInt4d($data, $pos)
    {
                $or24 = ord($data[$pos + 3]);
        if ($or24 >= 128) {
            $ord24 = -abs((256 - $or24) << 24);
        } else {
            $ord24 = ($or24 & 127) << 24;
        }
        return ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | $ord24;
    }
}
