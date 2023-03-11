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

    private array $x = [];

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
        ksort($timeTable);
        dump($timeTable);
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

        foreach ($employers as $employer) {
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

        $timeTable[$date][$employer['name']] = $preference;

        if ($preference === TimeTableRule2023::WORKING_ON_THE_NIGHT) {
            $timeTable[$startDate->copy()->addDay()->format('Y-m-d')][$employer['name']]
                = TimeTableRule2023::FREE_DAY_AFTER_NIGHT_TURN;
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
            && !$this->hasMaximumNightsTurns($timeTable, $employerName)
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
            && !$this->hasMaximumNightsTurns($timeTable, $employer['name'])
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

    private function hasMaximumNightsTurns(array $timeTable, string $employerName): bool
    {
        $noOfTurns = 0;

        foreach ($timeTable as $employer) {
            foreach ($employer as $name => $option) {
                if ($name !== $employerName) {
                    continue;
                }

                if ($option === TimeTableRule2023::WORKING_ON_THE_NIGHT) {
                    $noOfTurns++;
                }
            }
        }

        return $noOfTurns >= TimeTableRule2023::MAX_NO_OF_NIGHTS;
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

        if (!isset($this->x[$employerName]) || $this->x[$employerName] < $noOfTurns) {
            $this->x[$employerName] = $noOfTurns;
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
}
