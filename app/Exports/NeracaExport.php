<?php

namespace App\Exports;

use Illuminate\Contracts\View\View;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\Exportable;

use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\BeforeExport;
use Maatwebsite\Excel\Events\AfterSheet;

use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;


class NeracaExport implements FromView,  WithStyles, WithEvents
{
    private $assetArray;
    private $equityArray;
    private $liabilityArray;
    private $years;
    private $year;
    private $equitas;
    private $company;

    use Exportable;

    public function __construct(
        $assetArray, 
        $equityArray,
        $liabilityArray,
        $years,
        $year,
        $company,
        $equitas    
        ) 
    {
        $this->assetArray = $assetArray;
        $this->equityArray = $equityArray;
        $this->liabilityArray = $liabilityArray;
        $this->years = $years;
        $this->year = $year;
        $this->equitas = $equitas;
        $this->company = $company;
    }

    public function view(): View
    {
        
        return view('user.neracaExportExcel', [
            'company' => $this->company,
            'assetArray' => $this->assetArray,
            'equityArray' => $this->equityArray,
            'liabilityArray' => $this->liabilityArray,
            'years' => $this->years,
            'year' => $this->year,
            'equitas' => $this->equitas,
        ]);
    }

    public function styles(Worksheet $sheet)
    {
        return [
            // Style the first row as bold text.
            1 => [
                'font' => [
                    'bold' => true,
                    'size' => 18
                    ]],
            2 => [
                'font' => [
                   'bold' => true,
                   'size' => 18
                   ]], 

            4 => [
                'font' => [
                    'bold' => true,
                    'size' => 18
                    ]],

            // Styling a specific cell by coordinate.
            3 => [
                'font' => [
                    'italic' => true,
                    'size' => 16
                    ]],


            // Styling an entire column.
            'C'  => ['font' => ['size' => 16]],
            
            'A28' => [
                'font' => [
                    'bold' => true,
                    'size' => 16
                    ]],

            'A32' => [
                'font' => [
                    'bold' => true,
                    'size' => 16
                    ]],
        ];
    }
    public function registerEvents(): array
    {
        $styleArray1 = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ],
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK
                ],
            ],
        ];

        $styleArray2 = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ],
            ],
        ];

        $styleArray3 = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK
                ]
            ],
        ];
    

        return [
            AfterSheet::class    => function(AfterSheet $event) use (
                $styleArray1, 
                $styleArray2,
                $styleArray3 ) 
                
                {
                
                $event->sheet->getStyle('A4:B4')->ApplyFromArray($styleArray3);
                $event->sheet->getStyle('A5:B26')->ApplyFromArray($styleArray1);
                $event->sheet->getStyle('A25:B25')->ApplyFromArray($styleArray3);
                $event->sheet->getStyle('A26:B26')->ApplyFromArray($styleArray3);
                $event->sheet->getStyle('B5:B24')->ApplyFromArray($styleArray1);

                $event->sheet->getStyle('A28:A29')->ApplyFromArray($styleArray3);
                $event->sheet->getStyle('A32:A33')->ApplyFromArray($styleArray3);
            },
        ];
    }
}
