import { SharePointQueryable, SharePointQueryableInstance, SharePointQueryableCollection } from "./sharepointqueryable";

/**
 * Describes regional settings ODada object
 */
export class RegionalSettings extends SharePointQueryableInstance {

    /**
     * Creates a new instance of the RegionalSettings class
     *
     * @param baseUrl The url or SharePointQueryable which forms the parent of this regional settings collection
     */

    constructor(baseUrl: string | SharePointQueryable, path = "regionalsettings") {
        super(baseUrl, path);
    }

    /**
     * Gets installed languages
     */
    public get installedLanguages(): InstalledLanguages {
        return new InstalledLanguages(this);
    }

    /**
     * Gets time zone
     */
    public get timeZone(): TimeZone {
        return new TimeZone(this);
    }

    /**
     * Gets time zones
     */
    public get timeZones(): TimeZones {
        return new TimeZones(this);
    }

}

export interface RegionalSettingsProps {
    AdjustHijriDays: number;
    AlternateCalendarType: number;
    AM: string;
    CalendarType: number;
    Collation: number;
    CollationLCID: number;
    DateFormat: number;
    DateSeparator: string;
    DecimalSeparator: string;
    DigitGrouping: string;
    FirstDayOfWeek: number;
    FirstWeekOfYear: number;
    IsEastAsia: boolean;
    IsRightToLeft: boolean;
    IsUIRightToLeft: boolean;
    ListSeparator: string;
    LocaleId: number;
    NegativeSign: string;
    NegNumberMode: number;
    PM: string;
    PositiveSign: string;
    ShowWeeks: boolean;
    ThousandSeparator: string;
    Time24: boolean;
    TimeMarkerPosition: number;
    TimeSeparator: string;
    WorkDayEndHour: number;
    WorkDays: number;
    WorkDayStartHour: number;
}

/**
 * Describes installed languages ODada queriable collection
 */
export class InstalledLanguages extends SharePointQueryableInstance {
    constructor(baseUrl: string | SharePointQueryable, path = "installedlanguages") {
        super(baseUrl, path);
    }
}

/**
 * Describes TimeZone ODada object
 */
export class TimeZone extends SharePointQueryableInstance {
    constructor(baseUrl: string | SharePointQueryable, path = "timezone") {
        super(baseUrl, path);
    }

    /**
     * Gets an Local Time by UTC Time
     *
     * @param utcTime UTC Time as Date or ISO String
     */
    public utcToLocalTime(utcTime: string | Date): Promise<{ localTime: string }> {
        let dateIsoString: string;
        if (typeof utcTime === "string") {
            dateIsoString = utcTime;
        } else {
            dateIsoString = utcTime.toISOString();
        }
        return this.clone(TimeZone, `utctolocaltime(@date)?@date='${dateIsoString}'`)
            .postCore()
            .then(res => {
                return {
                    localTime: res.UTCToLocalTime,
                };
            });
    }

    /**
     * Gets an UTC Time by Local Time
     *
     * @param localTime Local Time as Date or ISO String
     */
    public localTimeToUTC(localTime: string | Date): Promise<{ utcTime: string }> {
        let dateIsoString: string;
        if (typeof localTime === "string") {
            dateIsoString = localTime;
        } else {
            dateIsoString = localTime.toISOString();
        }
        return this.clone(TimeZone, `localtimetoutc(@date)?@date='${dateIsoString}'`)
            .postCore()
            .then(res => {
                return {
                    utcTime: res.LocalTimeToUTC,
                };
            });
    }
}

/**
 * Describes time zones queriable collection
 */
export class TimeZones extends SharePointQueryableCollection {
    constructor(baseUrl: string | SharePointQueryable, path = "timezones") {
        super(baseUrl, path);
    }

    // https://msdn.microsoft.com/en-us/library/office/jj247008.aspx - timezones ids
    /**
     * Gets an TimeZone by id
     *
     * @param id The integer id of the timezone to retrieve
     */
    public getById(id: number): TimeZone {
        const tz: TimeZone = new TimeZone("", `${this.parentUrl}/timezones(${id})`);
        tz.get = (...args: any[]) => {
            return (tz as any).postCore(args[1], args[0]);
        };
        // Redefining get to trigger POST reqeust as
        // `/timezones(${id})` only supports 'POST' method
        return tz;
    }

}
