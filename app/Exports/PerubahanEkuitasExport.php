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


class PerubahanEkuitasExport implements FromView,  WithStyles, WithEvents
{
    private $equityArray;
    private $saldo_berjalan;
    private $company;

    use Exportable;

    public function __construct(
        $equityArray, 
        $saldo_berjalan,
        $company 
        ) 
    {
        $this->equityArray = $equityArray;
        $this->saldo_berjalan = $saldo_berjalan;
        $this->company = $company;
    }

    public function view(): View
    {
        
        return view('user.perubahanEkuitasExportExcel', [
            'equityArray' => $this->equityArray,
            'saldo_berjalan' => $this->saldo_berjalan,
            'company' => $this->company
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

            // Styling a specific cell by coordinate.
            3 => [
                'font' => [
                    'italic' => true,
                    'size' => 16
                    ]],


            // Styling an entire column.
            'C'  => ['font' => ['size' => 16]],
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
                
                $event->sheet->getStyle('A4:c4')->ApplyFromArray($styleArray3);

                $event->sheet->getStyle('A5:C7')->ApplyFromArray($styleArray1);

                $event->sheet->getStyle('A5:A7')->ApplyFromArray($styleArray1);

                $event->sheet->getStyle('B5:B7')->ApplyFromArray($styleArray1);

                $event->sheet->getStyle('A8:C9')->ApplyFromArray($styleArray3);

            },
        ];
    }
}
