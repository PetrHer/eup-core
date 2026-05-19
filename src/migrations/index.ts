import * as v001 from './v001';
import * as v002 from './v002';
import * as v003 from './v003';
import * as v004 from './v004';
import * as v005 from './v005';
import * as v006 from './v006';
import * as v007 from './v007';
import * as v008 from './v008';
import * as v009 from './v009';
import * as v010 from './v010';
import * as v011 from './v011';
import * as v012 from './v012';
import * as v013 from './v013';
import * as v014 from './v014';
import * as v015 from './v015';
import * as v016 from './v016';
import * as v017 from './v017';
import * as v018 from './v018';
import * as v019 from './v019';
import * as v020 from './v020';
import * as v021 from './v021';
import * as v022 from './v022';
import * as v023 from './v023';
import * as v024 from './v024';
import * as v025 from './v025';
import { IApiClient } from '../api/IApiClient';


const migrations = [v001, v002, v003, v004, v005, v006, v007, v008, v009, v010, v011, v012, v013, v014, v015, v016, v017, v018, v019, v020, v021, v022, v023, v024, v025];

export const runMigrations = async (currentVersion: number, apiClient: IApiClient, setInstallingVersion: (v: number) => void): Promise<boolean> => {
    for (const migration of migrations) {
        try {
            if (migration.version > currentVersion) {
                console.log(`Running migration v${String(migration.version).padStart(3, '0')}...`);
                setInstallingVersion(migration.version);
                const result = await migration.migrate(apiClient);
                if (!result) {
                    return false;
                }
            }
        } catch (e) {
            await apiClient.sp.log(`Migration v${migration.version.toString().padStart(3, '0')} failed.`, e);
            console.error(e);
        }
    }
    return true;
};

export const lastVersion = Math.max(...migrations.map(m => m.version));