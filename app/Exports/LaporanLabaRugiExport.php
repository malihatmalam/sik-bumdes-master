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


class LaporanLabaRugiExport implements FromView,  WithStyles, WithEvents
{
    private $incomeArray;
    private $expenseArray;
    private $years;
    private $year;
    private $othersIncomeArray;
    private $othersExpenseArray;
    private $income;
    private $expense;
    private $company;
    private $othersIncome;
    private $othersExpense;

    use Exportable;

    public function __construct(
        $incomeArray,
        $expenseArray, 
        $years,
        $year,
        $othersIncomeArray,
        $othersExpenseArray,
        $income, 
        $expense,
        $company,
        $othersIncome, 
        $othersExpense ) 
    {
        $this->incomeArray = $incomeArray;
        $this->expenseArray = $expenseArray;
        $this->years = $years;
        $this->year = $year;
        $this->othersIncomeArray = $othersIncomeArray;
        $this->othersExpenseArray = $othersExpenseArray;
        $this->income = $income;
        $this->expense = $expense;
        $this->company = $company;
        $this->othersIncome = $othersIncome;
        $this->othersExpense = $othersExpense;

    }

    public function view(): View
    {
        
        return view('user.laporanLabaRugiExportExcel', [
            'company_name' => $this->company,
            'incomeArray' => $this->incomeArray,
            'expenseArray' => $this->expenseArray,
            'years' => $this->years,
            'year' => $this->year,
            'othersIncomeArray' => $this->othersIncomeArray,
            'othersExpenseArray' => $this->othersExpenseArray,
            'income' => $this->income,
            'expense' => $this->expense,
            'othersIncome' => $this->othersIncome,
            'othersExpense' => $this->othersExpense
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

            5 => [
                'font' => [
                    'bold' => true,
                    'size' => 14
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
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK
                ],
            ],
        ];
    

        return [
            AfterSheet::class    => function(AfterSheet $event) use (
                $styleArray1, 
                $styleArray2 ) 
                
                {

                $cellRange = 'A1:W1'; // All headers
                $event->sheet->getDelegate()->getStyle($cellRange)->getFont()->setSize(20);
                
                $event->sheet->getStyle('A6:B18')->ApplyFromArray($styleArray1);

                $event->sheet->getStyle('A5:B5')->ApplyFromArray($styleArray2);

                $event->sheet->getStyle('A18:B18')->ApplyFromArray($styleArray1);
                $event->sheet->getStyle('A18')->ApplyFromArray($styleArray1); 

                $event->sheet->getStyle('A6:A17')->ApplyFromArray($styleArray1);                
            },
        ];
    }
}
