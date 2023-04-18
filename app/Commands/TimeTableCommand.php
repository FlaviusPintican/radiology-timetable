<?php

declare(strict_types=1);

namespace App\Commands;

use App\Helpers\TimeTableHelper;
use App\Rules\TimeTableRule2023;
use Carbon\Carbon;
use Exception;
use InvalidArgumentException;
use LaravelZero\Framework\Commands\Command;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Throwable;

class TimeTableCommand extends Command
{
    /**
     * The signature of the command.
     *
     * @var string
     */
    protected $signature = 'radiology:timetable';

    /**
     * The description of the command.
     *
     * @var string
     */
    protected $description = 'Generate radiology timetable';

    private ?string $filePath = null;

    private int $noOfWorkingPersonsOnWeekends = 0;

    private const TABLE_MAP_COLUMNS = [
        'Nume' => 'name',
        'Interval concediu' => 'holiday_interval',
        'Preferinte program' => 'favorite_schedule_time',
        'A muncit noaptea ultimei zile din luna trecuta' => 'last_working_month_night',
        'Lucreaza de noaptea' => 'working_on_night',
        'Lucreaza in weekend' => 'working_on_weekend',
    ];

    /**
     * @throws Exception
     */
    public function handle(): void
    {
        try {
            $this->validateExcelFile();
        } catch (Throwable $exception) {
            $this->output->error($exception->getMessage());

            return;
        }

        $this->generateEmployersTimeTable();
    }

    /**
     * @throws Exception
     */
    private function validateExcelFile(): void
    {
        if (
            !is_file($this->getFilePath())
            &&
            (
                !str_ends_with('.xslx', $this->getFilePath())
                || !str_ends_with('.xslm', $this->getFilePath())
                || !str_ends_with('.xslb', $this->getFilePath())
                || !str_ends_with('.xltx', $this->getFilePath())
                || !str_ends_with('.xls', $this->getFilePath())
            )
        ) {
            throw new InvalidArgumentException('Fisierul nu a fost gasit sau nu este un excel valid');
        }
    }

    private function transformExcelIntoArray(): array
    {
        $spreadsheet = IOFactory::load($this->getFilePath());

        return $spreadsheet->getActiveSheet()->toArray();
    }

    /**
     * @return array<Carbon>
     */
    private function getDateInterval(): array
    {
        $now = Carbon::now();

        if ($now->day > 1) {
            $now->addMonth();
        }

        return [
            $now->clone()->startOfMonth(),
            $now->clone()->endOfMonth(),
        ];
    }

    /**
     * @throws Exception
     */
    private function generateEmployersTimeTable(): void
    {
        $employers = $this->getFormattedEmployers();
        $timeTable = [];
        $favoriteScheduleTimes = [];
        $preference = null;

        $this->applyPreferencesAndHolidays($employers, $timeTable, $favoriteScheduleTimes, $preference);
        $this->addAllNights($employers, $timeTable, $favoriteScheduleTimes);
        $this->addExtraNights($employers, $timeTable, $favoriteScheduleTimes);
        $this->addAllWeekendsTurns($employers, $timeTable, $favoriteScheduleTimes);
        $this->addPublicHolidaysTurns($employers, $timeTable, $favoriteScheduleTimes);
        $this->addFreeDaysForEveryEmployers($employers, $timeTable);
        $this->addMinimumMornings($employers, $timeTable);
        $this->addMinimumAfternoons($employers, $timeTable);
        $this->addMorningsAndAfternoonsForSpecificEmployers($employers, $timeTable);
        $this->addAllMinimumMornings($employers, $timeTable);
        $this->addAllMinimumAfternoons($employers, $timeTable);
        $this->addAllMaximumMornings($employers, $timeTable);
        $this->addAllMaximumAfternoons($employers, $timeTable);

        $this->saveTimeTable($employers, $timeTable);
    }

    /**
     * @throws Exception
     */
    private function getFormattedEmployers(): array
    {
        $employers = $this->transformExcelIntoArray();

        if (count($employers) === 0) {
            throw new InvalidArgumentException('Excelul este gol.');
        }

        $employerHeaders = array_shift($employers);
        $formattedEmployers = [];
        $priorityEmployer = [];

        foreach ($employers as $key => $employer) {
            $formattedEmployer = [];

            foreach ($employer as $index => $value) {
                if (self::TABLE_MAP_COLUMNS[$employerHeaders[$index]] === 'name' && !$value) {
                    continue;
                }

                $formattedEmployer[self::TABLE_MAP_COLUMNS[$employerHeaders[$index]]] = $value;
            }

            if (TimeTableRule2023::WORKING_ON_WEEKENDS[($formattedEmployer['working_on_weekend'] ?? 'DA')]) {
                $this->noOfWorkingPersonsOnWeekends++;
            }

            $formattedEmployer['order_index'] = $key;

            if (($formattedEmployer['name'] ?? '') === TimeTableRule2023::PRIORITY_EMPLOYER) {
                $priorityEmployer = $formattedEmployer;
            } else {
                $formattedEmployers[] = $formattedEmployer;
            }
        }

        shuffle($formattedEmployers);

        if ($priorityEmployer) {
            array_unshift($formattedEmployers, $priorityEmployer);
        }

        return $formattedEmployers;
    }

    /**
     * Check if it's a holiday paid
     *
     * @param Carbon $date
     * @param string $holidayIntervals
     * @return bool
     */
    private function isHoliday(Carbon $date, string $holidayIntervals): bool
    {
        foreach (TimeTableHelper::removeEmptyValuesFromDateInterval($holidayIntervals) as $holidayInterval) {
            [$startDate, $endDate] = explode(':', $holidayInterval, 2);

            if (
                $date->getTimestamp() >= Carbon::parse($startDate)->getTimestamp()
                && $date->getTimestamp() <= Carbon::parse($endDate)->getTimestamp()
            ) {
                return true;
            }
        }

        return false;
    }

    private function getPreferencesForSpecificDate(string $preferences): ?array
    {
        $preferences = TimeTableHelper::removeEmptyValuesFromDateInterval($preferences);

        if (count($preferences) === 0) {
            return null;
        }

        $formattedPreferences = [];
        $firstDayOfMonth = Carbon::now()->addMonth()->startOfMonth();
        $lastDayOfMonth = Carbon::now()->addMonth()->endOfMonth()->setTime(0, 0);

        foreach ($preferences as $preference) {
            [$preferenceDate, $options] = explode('(', $preference, 2);
            $options = trim(substr($options, 0, -1));
            [$startDate, $endDate] = explode(':', $preferenceDate, 2) + ['', null];
            $endDate = $endDate ?? $startDate;

            if ($this->isValidDateInterval($startDate, $endDate)) {
                $newStartDate = Carbon::parse($startDate);
                $newEndDate = Carbon::parse($endDate);

                while ($newStartDate <= $newEndDate) {
                    $formattedPreferences[$newStartDate->format('Y-m-d')] = $options;
                    $newStartDate->addDay();
                }
            } else {
                if (str_contains($startDate, '|')) {
                    $dates = array_filter(
                        explode('|', $startDate),
                        fn(mixed $value) => !in_array($value, ['', null])
                    );
                } else {
                    $dates = array_filter(
                        explode('-', $startDate),
                        fn(mixed $value) => !in_array($value, ['', null])
                    );

                    $newDates = [];

                    for ($index = $dates[0] ?? 1; $index <= $dates[1] ?? 5; $index++) {
                        $newDates[] = $index;
                    }

                    $dates = $newDates;
                }

                $startDate = $firstDayOfMonth->clone();

                while ($startDate < $lastDayOfMonth) {
                    /** @var Carbon $startDate */
                    $startDate = $startDate->clone()->startOfWeek(0);

                    foreach ($dates as $date) {
                        $date = (int)$date;

                        if (
                            $startDate->getTimestamp() < $firstDayOfMonth->getTimestamp()
                            && $startDate->clone()->addDays($date)->month < $firstDayOfMonth->month
                        ) {
                            continue;
                        }

                        $formattedPreferences[$startDate->clone()->addDays($date)->format('Y-m-d')] = $options;

                        if ($startDate >= $lastDayOfMonth) {
                            break;
                        }
                    }

                    $startDate = $firstDayOfMonth->addWeek();
                }
            }
        }

        return $formattedPreferences;
    }

    private function getFilePath(): string
    {
        if (null === $this->filePath) {
            $this->filePath = storage_path('app') . '/timetable-input.xlsx';
        }

        return $this->filePath;
    }

    private function getOutputFilePath(): string
    {
        return storage_path('app') . '/timetable-output.xlsx';
    }

    private function getPreference(array &$favoriteScheduleTimes, array $employer, string $date): ?string
    {
        $preferences = $this->getPreferencesForSpecificDate(
            $employer['favorite_schedule_time'] ?? ''
        )[$date] ?? null;

        if (null !== $preferences) {
            $favoriteScheduleTimes[$date][$employer['name']]['preferences'] = $preferences;
        }

        $preference = $favoriteScheduleTimes[$date][$employer['name']]['preferences'] ?? '';

        return match ($preference) {
            TimeTableRule2023::FREE_DAY => TimeTableRule2023::FREE_DAY,
            TimeTableRule2023::WORKING_ON_THE_MORNING => TimeTableRule2023::WORKING_ON_THE_MORNING,
            TimeTableRule2023::WORKING_ON_THE_AFTERNOON => TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
            TimeTableRule2023::WORKING_ON_THE_NIGHT => TimeTableRule2023::WORKING_ON_THE_NIGHT,
            TimeTableRule2023::FREE_NATIONAL_DAY => TimeTableRule2023::FREE_NATIONAL_DAY,
            default => null
        };
    }

    private function isValidDateInterval(string $startDate, string $endDate): bool
    {
        try {
            Carbon::parse($startDate);
            Carbon::parse($endDate);
            $isValidDateInterval = true;
        } catch (Throwable) {
            $isValidDateInterval = false;
        }

        return $isValidDateInterval;
    }

    private function applyWorkingLastWorkingMonthNight(array &$timeTable, array $employer, string $date): void
    {
        if (
            (TimeTableRule2023::WORKING_LAST_MONTH_NIGHT[$employer['last_working_month_night'] ?? 'NU'])
            && Carbon::parse($date)->day === 1
        ) {
            $timeTable[$date][$employer['name']] = TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN;
        }
    }

    private function applyHolidayForSpecificDate(
        Carbon $startDate,
        array &$timeTable,
        array $employer,
        string $date
    ): void {
        if ($this->isHoliday($startDate, $employer['holiday_interval'] ?? '')) {
            $timeTable[$date][$employer['name']] = TimeTableRule2023::HOLIDAY;
        }
    }

    private function applyPreference(
        Carbon $startDate,
        array &$timeTable,
        array $employer,
        string $date,
        ?string $preference
    ): void {
        if (null === $preference) {
            return;
        }

        if (!isset($timeTable[$date][$employer['name']])) {
            $timeTable[$date][$employer['name']] = $preference;
        }

        if ($preference === TimeTableRule2023::WORKING_ON_THE_NIGHT) {
            $nextDay = $startDate->copy()->addDay()->format('Y-m-d');

            if (!isset($timeTable[$nextDay][$employer['name']])) {
                $timeTable[$nextDay][$employer['name']] = TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN;
            }
        }
    }

    private function applyNonWeekendNights(
        Carbon $startDate,
        array &$timeTable,
        array $favoriteScheduleTimes,
        string $employerName,
        string $date,
        int $maximumTurns
    ): void {
        if (
            !$startDate->isWeekend()
            && !$startDate->isFriday()
            && !in_array($date, TimeTableRule2023::$publicHolidays, true)
            && !$this->hasMaximumTurns(
                $timeTable,
                $employerName,
                TimeTableRule2023::WORKING_ON_THE_NIGHT,
                TimeTableRule2023::MAX_NO_OF_NIGHTS
            )
            && $this->canHaveNightTurn(
                $timeTable,
                $date,
                $maximumTurns
            )
            && !isset($timeTable[$startDate->copy()->addDay()->format('Y-m-d')][$employerName])
            && !isset(
                $favoriteScheduleTimes[$startDate->copy()->addDay()->format('Y-m-d')][$employerName]['preferences']
            )
        ) {
            $timeTable[$date][$employerName] = TimeTableRule2023::WORKING_ON_THE_NIGHT;

            if ($startDate->copy()->addDay()->month === $startDate->month) {
                $timeTable[$startDate->copy()->addDay()->format('Y-m-d')][$employerName]
                    = TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN;
            }
        }
    }

    private function applyAllNights(
        Carbon $startDate,
        array &$timeTable,
        array $favoriteScheduleTimes,
        array $employer,
        string $date,
    ): void {
        $preference = $favoriteScheduleTimes[$date][$employer['name']]['preferences'] ?? null;

        if ($this->hasPreferences($employer, $preference)) {
            return;
        }

        if (
            (
                $startDate->isWeekend()
                || $startDate->isFriday()
                || in_array($date, TimeTableRule2023::$publicHolidays, true)
            )
            && !$this->hasMaximumTurns(
                $timeTable,
                $employer['name'],
                TimeTableRule2023::WORKING_ON_THE_NIGHT,
                TimeTableRule2023::MAX_NO_OF_NIGHTS,
            )
            && $this->canHaveNightTurn(
                $timeTable,
                $date,
                TimeTableRule2023::MAX_NO_OF_PERSONS_FROM_ONE_TURN_WEEKEND
            )
            && !isset($timeTable[$startDate->copy()->addDay()->format('Y-m-d')][$employer['name']])
            && !isset(
                $favoriteScheduleTimes[$startDate->copy()->addDay()->format('Y-m-d')][$employer['name']]['preferences']
            )
        ) {
            $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_NIGHT;

            if ($startDate->copy()->addDay()->month === $startDate->month) {
                $timeTable[$startDate->copy()->addDay()->format('Y-m-d')][$employer['name']]
                    = TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN;
            }
        }

        $this->applyNonWeekendNights(
            $startDate,
            $timeTable,
            $favoriteScheduleTimes,
            $employer['name'],
            $date,
            TimeTableRule2023::MAX_NO_OF_PERSONS_FROM_ONE_TURN_WEEKEND
        );
    }

    private function applyNights(
        Carbon $startDate,
        array &$timeTable,
        array $favoriteScheduleTimes,
        array $employer,
        string $date,
    ): void {
        $preference = $favoriteScheduleTimes[$date][$employer['name']]['preferences'] ?? null;

        if ($this->hasPreferences($employer, $preference)) {
            return;
        }

        $this->applyNonWeekendNights(
            $startDate,
            $timeTable,
            $favoriteScheduleTimes,
            $employer['name'],
            $date,
            TimeTableRule2023::MAX_NO_OF_PERSONS_FROM_MONDAY_TO_THURSDAY_NIGHTS
        );
    }

    private function canHaveNightTurn(array $timeTable, string $date, int $maximumTurns): bool
    {
        $noOfTurns = 0;

        foreach ($timeTable[$date] ?? [] as $option) {
            if ($option === TimeTableRule2023::WORKING_ON_THE_NIGHT) {
                $noOfTurns++;
            }
        }

        return $noOfTurns < $maximumTurns;
    }

    private function hasMaximumTurns(
        array $timeTable,
        string $employerName,
        string $employerOption,
        int $noOfTurns
    ): bool {
        $totalTurns = 0;

        foreach ($timeTable as $employer) {
            foreach ($employer as $name => $option) {
                if ($name !== $employerName) {
                    continue;
                }

                if ($option === $employerOption) {
                    $totalTurns++;
                }
            }
        }

        return $totalTurns >= $noOfTurns;
    }

    private function hasPreferences(array $employer, ?string $preference): bool
    {
       return $preference
           || $preference === TimeTableRule2023::FREE_ON_WEEKEND_DAY
           || $preference === TimeTableRule2023::WORKING_ON_THE_MORNING_OR_ON_THE_AFTERNOON
           || (strtoupper($employer['working_on_night'] ?? 'DA') === 'NU')
           || (strtoupper($employer['working_on_weekend'] ?? 'DA') !== 'DA');
    }

    private function applyPreferencesAndHolidays(
        array $employers,
        array &$timeTable,
        array &$favoriteScheduleTimes,
        ?string &$preference
    ): void {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                $this->applyWorkingLastWorkingMonthNight($timeTable, $employer, $date);
                $this->applyHolidayForSpecificDate($startDate, $timeTable, $employer, $date);
                $preference = $this->getPreference($favoriteScheduleTimes, $employer, $date);
                $this->applyPreference($startDate, $timeTable, $employer, $date, $preference);
            }

            $startDate->addDay();
        }
    }

    private function addExtraNights(array $employers, array &$timeTable, array $favoriteScheduleTimes): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                $this->applyNights($startDate, $timeTable, $favoriteScheduleTimes, $employer, $date);
            }

            $startDate->addDay();
        }
    }

    private function addAllNights(array $employers, array &$timeTable, array $favoriteScheduleTimes): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                $this->applyAllNights($startDate, $timeTable, $favoriteScheduleTimes, $employer, $date);
            }

            $startDate->addDay();
        }
    }

    private function getTotalTurnsOnWeekend(): int
    {
        [$startDate, $endDate] = $this->getDateInterval();
        $total = 0;

        while ($startDate <= $endDate) {
            if ($startDate->isFriday()) {
                $total++;
            }

            if ($startDate->isSaturday()) {
                $total += 4;
            }

            if ($startDate->isSunday()) {
                $total += 3;
            }

            $startDate->addDay();
        }

        return $total;
    }

    private function getNumberOfMediumWeekendWorkingHours(): int
    {
        $turns = (int)round($this->getTotalTurnsOnWeekend() / $this->noOfWorkingPersonsOnWeekends);

        if (
            ($turns * TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS)
            < TimeTableRule2023::MIN_NO_OF_WEEKEND_WORKING_HOURS
        ) {
            $turns = TimeTableRule2023::MIN_NO_OF_WEEKEND_WORKING_HOURS
                / TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS;
        }

        if (
            ($turns * TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS)
            > TimeTableRule2023::MAX_NO_OF_WEEKEND_WORKING_HOURS
        ) {
            $turns = TimeTableRule2023::MAX_NO_OF_WEEKEND_WORKING_HOURS
                / TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS;
        }

        return $turns * TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS;
    }

    private function addWeekendsTurns(
        array $employers,
        array &$timeTable,
        array $favoriteScheduleTimes,
        int $noOfWorkingHours
    ): void {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                if (!$startDate->isWeekend() && !$startDate->isFriday()) {
                    continue;
                }

                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                $this->applyWeekendsTurns($timeTable, $favoriteScheduleTimes, $employer, $date, $noOfWorkingHours);
            }

            $startDate->addDay();
        }
    }

    private function applyWeekendsTurns(
        array &$timeTable,
        array $favoriteScheduleTimes,
        array $employer,
        string $date,
        int $noOfWorkingHours
    ): void {
        $preference = $favoriteScheduleTimes[$date][$employer['name']]['preferences'] ?? null;

        if ($this->hasPreferences($employer, $preference)) {
            return;
        }

        if (Carbon::parse($date)->isFriday()) {
            return;
        }

        if ($this->hasMinimumWeekendTurns($timeTable, $employer['name'], $noOfWorkingHours)) {
            return;
        }

        $weekendOptions = $this->getWeekendTurnOptions($timeTable)[$date] ?? [];

        if (!in_array(TimeTableRule2023::WORKING_ON_THE_MORNING, $weekendOptions, true)) {
            $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_MORNING;

            return;
        }

        if (!in_array(TimeTableRule2023::WORKING_ON_THE_AFTERNOON, $weekendOptions, true)) {
            $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_AFTERNOON;
        }
    }

    private function hasMinimumWeekendTurns(array $timeTable, string $employerName, int $noOfWorkingHours): bool
    {
        $noOfTurns = 0;

        foreach ($timeTable as $date => $employer) {
            $dateObject = Carbon::parse($date);

            if (!$dateObject->isWeekend() && !$dateObject->isFriday()) {
                continue;
            }

            foreach ($employer as $name => $option) {
                if ($name !== $employerName) {
                    continue;
                }

                if ($dateObject->isSaturday() && $option === TimeTableRule2023::WORKING_ON_THE_NIGHT) {
                    $noOfTurns += 2;
                    continue;
                }

                if (
                    $dateObject->isSaturday()
                    && (
                        $option === TimeTableRule2023::WORKING_ON_THE_MORNING
                        || $option === TimeTableRule2023::WORKING_ON_THE_AFTERNOON
                    )
                ) {
                    $noOfTurns++;
                    continue;
                }

                if (
                    $dateObject->isSunday()
                    && (
                        $option === TimeTableRule2023::WORKING_ON_THE_MORNING
                        || $option === TimeTableRule2023::WORKING_ON_THE_AFTERNOON
                        || $option === TimeTableRule2023::WORKING_ON_THE_NIGHT
                    )
                ) {
                    $noOfTurns++;
                    continue;
                }

                if ($dateObject->isFriday() && $option === TimeTableRule2023::WORKING_ON_THE_NIGHT) {
                    $noOfTurns++;
                }
            }
        }

        return $noOfTurns >= ($noOfWorkingHours / TimeTableRule2023::MAX_NO_OF_DAILY_WORKING_HOURS);
    }

    private function getWeekendTurnOptions(array $timeTable): array
    {
        $dates = [];

        foreach ($timeTable as $date => $employer) {
            $dateObject = Carbon::parse($date);

            if (!$dateObject->isWeekend()) {
                continue;
            }

            foreach ($employer as $option) {
                $dates[$date][] = $option;
            }
        }

        return $dates;
    }

    private function getTurnOptionsForSpecificDate(array $timeTable, string $date): array
    {
        $options = [];

        foreach ($timeTable[$date] ?? [] as $option) {
            $options[] = $option;
        }

        return $options;
    }

    private function addAllWeekendsTurns(array $employers, array &$timeTable, array $favoriteScheduleTimes): void
    {
        $this->addWeekendsTurns(
            $employers,
            $timeTable,
            $favoriteScheduleTimes,
            TimeTableRule2023::MIN_NO_OF_WEEKEND_WORKING_HOURS
        );
        $this->addWeekendsTurns(
            $employers,
            $timeTable,
            $favoriteScheduleTimes,
            $this->getNumberOfMediumWeekendWorkingHours()
        );
        $this->addWeekendsTurns(
            $employers,
            $timeTable,
            $favoriteScheduleTimes,
            TimeTableRule2023::MAX_NO_OF_WEEKEND_WORKING_HOURS
        );
    }

    private function addPublicHolidaysTurns(array $employers, array &$timeTable, array $favoriteScheduleTimes): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (!in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                $preference = $favoriteScheduleTimes[$date][$employer['name']]['preferences'] ?? null;

                if ($this->hasPreferences($employer, $preference)) {
                    continue;
                }

                if (
                    !in_array(
                        TimeTableRule2023::WORKING_ON_THE_MORNING,
                        $this->getTurnOptionsForSpecificDate($timeTable, $date),
                        true
                    )
                ) {
                    $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_MORNING;
                }

                if (
                    !in_array(
                        TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
                        $this->getTurnOptionsForSpecificDate($timeTable, $date),
                        true
                    )
                ) {
                    $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_AFTERNOON;
                }
            }

            $startDate->addDay();
        }
    }

    private function getNoOfFreeDaysForEveryEmployers(array $timeTable): array
    {
        $employerDayNumbers = [];

        foreach ($timeTable as $date => $employer) {
            foreach ($employer as $name => $option) {
                if (!isset($employerDayNumbers[$name])) {
                    $employerDayNumbers[$name] = 0;
                }

                $dateObject = Carbon::parse($date);

                if (
                    in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && !$dateObject->isWeekend()
                    && !$dateObject->isFriday()
                    && in_array(
                        $option,
                        [
                            TimeTableRule2023::WORKING_ON_THE_MORNING,
                            TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
                            TimeTableRule2023::WORKING_ON_THE_NIGHT
                        ],
                        true
                    )
                ) {
                    $employerDayNumbers[$name]++;
                    continue;
                }

                if (
                    in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && $dateObject->isFriday()
                    && in_array(
                        $option,
                        [
                            TimeTableRule2023::WORKING_ON_THE_MORNING,
                            TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
                        ],
                        true
                    )
                ) {
                    $employerDayNumbers[$name]++;
                    continue;
                }

                if (
                    in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && $dateObject->isFriday()
                    && $option === TimeTableRule2023::WORKING_ON_THE_NIGHT
                ) {
                    $employerDayNumbers[$name] += 2;
                    continue;
                }

                if (
                    in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && $dateObject->isWeekend()
                    && in_array(
                        $option,
                        [
                            TimeTableRule2023::WORKING_ON_THE_MORNING,
                            TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
                            TimeTableRule2023::WORKING_ON_THE_NIGHT
                        ],
                        true
                    )
                ) {
                    $employerDayNumbers[$name] += 2;
                    continue;
                }

                if (
                    !in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && ($dateObject->isFriday() || $dateObject->isSunday())
                    && $option === TimeTableRule2023::WORKING_ON_THE_NIGHT
                ) {
                    $employerDayNumbers[$name]++;
                    continue;
                }

                if (
                    !in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && $dateObject->isWeekend()
                    && in_array(
                        $option,
                        [
                            TimeTableRule2023::WORKING_ON_THE_MORNING,
                            TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
                        ],
                        true
                    )
                ) {
                    $employerDayNumbers[$name]++;
                    continue;
                }

                if (
                    !in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && $dateObject->isSaturday()
                    && $option === TimeTableRule2023::WORKING_ON_THE_NIGHT
                ) {
                    $employerDayNumbers[$name] += 2;
                }
            }
        }

        return $employerDayNumbers;
    }

    private function addFreeDaysForEveryEmployers(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();
        $employersFreeDays = $this->getNoOfFreeDaysForEveryEmployers($timeTable);

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if ($startDate->isWeekend()) {
                    continue;
                }

                $noOfFreeDays = array_count_values(array_values($timeTable[$date]))[TimeTableRule2023::FREE_DAY] ?? 0;

                if (($employersFreeDays[$employer['name']] ?? 0) && $noOfFreeDays < 2) {
                    $timeTable[$date][$employer['name']] = TimeTableRule2023::FREE_DAY;
                    $employersFreeDays[$employer['name']]--;
                }
            }

            $startDate->addDay();
        }
    }

    private function getNumberOfWorkingHoursForSpecificWeek(array $timeTable, Carbon $date, string $employerName): int
    {
        $startOfWeek = $date->clone()->startOfWeek(1);

        if ($startOfWeek->month < $date->month) {
            $startOfWeek = $date;
        }

        $endOfWeek = $date->clone()->startOfWeek(0)->addWeek();
        $noOfHours = 0;

        foreach ($timeTable as $date => $employers) {
            $dateObject = Carbon::parse($date);

            if (
                $dateObject->getTimestamp() < $startOfWeek->getTimestamp()
                || $dateObject->getTimestamp() > $endOfWeek->getTimestamp()
            ) {
                continue;
            }

            foreach ($employers as $name => $option) {
                if ($name !== $employerName) {
                    continue;
                }

                $noOfHours += TimeTableHelper::mapEmployersHours()[$option] ?? 0;
            }
        }

        return $noOfHours;
    }

    private function addMinimumMornings(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if (strtoupper($employer['working_on_weekend'] ?? 'DA') !== 'DA') {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                $noOfMornings = array_count_values(
                    array_values($timeTable[$date])
                )[TimeTableRule2023::WORKING_ON_THE_MORNING] ?? 0;

                if ($noOfMornings >= TimeTableRule2023::MIN_NO_OF_PERSONS_FOR_MORNING_TURN) {
                    continue;
                }

                if (
                    $this->getNumberOfWorkingHoursForSpecificWeek($timeTable, $startDate, $employer['name'])
                    >= TimeTableRule2023::MAX_NO_OF_WEEKLY_WORKING_HOURS
                ) {
                    continue;
                }

                if (
                    !$this->hasMaximumTurns(
                        $timeTable,
                        $employer['name'],
                        TimeTableRule2023::WORKING_ON_THE_MORNING,
                        TimeTableRule2023::MIN_NO_OF_MORNINGS,
                    )
                ) {
                    $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_MORNING;
                }
            }

            $startDate->addDay();
        }
    }

    private function addMinimumAfternoons(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if (strtoupper($employer['working_on_weekend'] ?? 'DA') !== 'DA') {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                $noOfAfternoons = array_count_values(
                    array_values($timeTable[$date])
                )[TimeTableRule2023::WORKING_ON_THE_AFTERNOON] ?? 0;

                if ($noOfAfternoons >= TimeTableRule2023::MIN_NO_OF_PERSONS_FOR_AFTERNOON_TURN) {
                    continue;
                }

                if (
                    $this->getNumberOfWorkingHoursForSpecificWeek($timeTable, $startDate, $employer['name'])
                    >= TimeTableRule2023::MAX_NO_OF_WEEKLY_WORKING_HOURS
                ) {
                    continue;
                }

                if (
                    !$this->hasMaximumTurns(
                        $timeTable,
                        $employer['name'],
                        TimeTableRule2023::WORKING_ON_THE_AFTERNOON,
                        TimeTableRule2023::MIN_NO_OF_AFTERNOONS,
                    )
                ) {
                    $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_AFTERNOON;
                }
            }

            $startDate->addDay();
        }
    }

    private function addMorningsAndAfternoonsForSpecificEmployers(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                if (strtoupper($employer['working_on_weekend'] ?? 'DA') === 'DA') {
                    continue;
                }

                $options = [
                    TimeTableRule2023::WORKING_ON_THE_MORNING,
                    TimeTableRule2023::WORKING_ON_THE_AFTERNOON
                ];

                $timeTable[$date][$employer['name']] = $options[array_rand($options)];
            }

            $startDate->addDay();
        }
    }

    private function addAllMinimumMornings(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                $noOfMornings = array_count_values(
                    array_values($timeTable[$date])
                )[TimeTableRule2023::WORKING_ON_THE_MORNING] ?? 0;

                if ($noOfMornings >= TimeTableRule2023::MIN_NO_OF_PERSONS_FOR_MORNING_TURN) {
                    continue;
                }

                if (
                    $this->getNumberOfWorkingHoursForSpecificWeek($timeTable, $startDate, $employer['name'])
                    >= TimeTableRule2023::MAX_NO_OF_WEEKLY_WORKING_HOURS
                ) {
                    continue;
                }

                $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_MORNING;
            }

            $startDate->addDay();
        }
    }

    private function addAllMinimumAfternoons(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                $noOfAfternoons = array_count_values(
                    array_values($timeTable[$date])
                )[TimeTableRule2023::WORKING_ON_THE_AFTERNOON] ?? 0;

                if ($noOfAfternoons >= TimeTableRule2023::MIN_NO_OF_PERSONS_FOR_AFTERNOON_TURN) {
                    continue;
                }

                if (
                    $this->getNumberOfWorkingHoursForSpecificWeek($timeTable, $startDate, $employer['name'])
                    >= TimeTableRule2023::MAX_NO_OF_WEEKLY_WORKING_HOURS
                ) {
                    continue;
                }

                $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_AFTERNOON;
            }

            $startDate->addDay();
        }
    }

    private function addAllMaximumMornings(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                $noOfMornings = array_count_values(
                    array_values($timeTable[$date])
                )[TimeTableRule2023::WORKING_ON_THE_MORNING] ?? 0;

                if ($noOfMornings >= TimeTableRule2023::MAX_NO_OF_PERSONS_FOR_MORNING_TURN) {
                    continue;
                }

                if (
                    $this->getNumberOfWorkingHoursForSpecificWeek($timeTable, $startDate, $employer['name'])
                    >= TimeTableRule2023::MAX_NO_OF_WEEKLY_WORKING_HOURS
                ) {
                    continue;
                }

                $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_MORNING;
            }

            $startDate->addDay();
        }
    }

    private function addAllMaximumAfternoons(array $employers, array &$timeTable): void
    {
        [$startDate, $endDate] = $this->getDateInterval();

        while ($startDate <= $endDate) {
            foreach ($employers as $employer) {
                $date = $startDate->format('Y-m-d');

                if (isset($timeTable[$date][$employer['name']])) {
                    continue;
                }

                if ($startDate->isWeekend() || in_array($date, TimeTableRule2023::$publicHolidays, true)) {
                    continue;
                }

                $noOfAfternoons = array_count_values(
                    array_values($timeTable[$date])
                )[TimeTableRule2023::WORKING_ON_THE_AFTERNOON] ?? 0;

                if ($noOfAfternoons >= TimeTableRule2023::MAX_NO_OF_PERSONS_FOR_AFTERNOON_TURN) {
                    continue;
                }

                if (
                    $this->getNumberOfWorkingHoursForSpecificWeek($timeTable, $startDate, $employer['name'])
                    >= TimeTableRule2023::MAX_NO_OF_WEEKLY_WORKING_HOURS
                ) {
                    continue;
                }

                $timeTable[$date][$employer['name']] = TimeTableRule2023::WORKING_ON_THE_AFTERNOON;
            }

            $startDate->addDay();
        }
    }

    /**
     * @throws Exception
     */
    public function saveTimeTable(array $employers, array $timeTable)
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $employers = $this->getOrderedEmployers($employers);
        $noOfEmployers = count($employers) + 1;

        foreach ($employers as $key => $employer) {
            $sheet->setCellValue('A' . ($key + 2), $employer['name']);
        }

        $index = 1;

        foreach ($timeTable as $date => $employersTable) {
            [, $month, $day] = explode('-', $date, 3) + ['', '', ''];
            $column = TimeTableHelper::mapColumnTimeTable()[$index];
            $sheet->setCellValue($column . '1', $day . '/' . $month);

            if (Carbon::parse($date)->isWeekend()) {
                $spreadsheet
                    ->getActiveSheet()
                    ->getStyle($column . '1:' . $column . $noOfEmployers)
                    ->getFill()
                    ->setFillType(Fill::FILL_SOLID)
                    ->getStartColor()
                    ->setARGB(Color::COLOR_YELLOW);
            }

            foreach ($employersTable as $name => $option) {
                if (
                    in_array(
                        $option,
                        [TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN, TimeTableRule2023::FREE_NATIONAL_DAY],
                        true
                    )
                ) {
                    continue;
                }

                if (
                    in_array($date, TimeTableRule2023::$publicHolidays, true)
                    && $option === TimeTableRule2023::FREE_DAY
                ) {
                    continue;
                }

                $employerIndex = array_values(
                    array_filter($employers, fn (array $employer) => $name === $employer['name'])
                )[0]['order_index'];
                $sheet->setCellValue($column . ($employerIndex + 2), $option);
            }

            $index++;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($this->getOutputFilePath());
    }

    private function getOrderedEmployers(array $employers): array
    {
        usort(
            $employers,
            fn (array $firstEmployer, array $secondEmployer)
                => $firstEmployer['order_index'] <=> $secondEmployer['order_index']
        );

        return $employers;
    }
}
