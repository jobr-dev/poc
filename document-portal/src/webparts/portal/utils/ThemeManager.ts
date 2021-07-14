export enum TeamsTheme {
    default,
    dark,
    contrast
}
export class ThemeManager {
    public static palette: any;
    public static applyTeamsTheme(teamsTheme: TeamsTheme) {
        let theme = {
            name: "Teams",
            palette: {}
        };
        switch (teamsTheme) {
            case TeamsTheme.dark: {

                theme.palette = {
                    themePrimary: '#9796e3',
                    themeLighterAlt: '#060609',
                    themeLighter: '#181824',
                    themeLight: '#2d2d44',
                    themeTertiary: '#5b5a88',
                    themeSecondary: '#8584c8',
                    themeDarkAlt: '#a19fe6',
                    themeDark: '#aeadea',
                    themeDarker: '#c2c2ef',
                    neutralLighterAlt: '#323130',
                    neutralLighter: '#31302f',
                    neutralLight: '#2f2e2d',
                    neutralQuaternaryAlt: '#2c2b2a',
                    neutralQuaternary: '#2a2928',
                    neutralTertiaryAlt: '#282726',
                    neutralTertiary: '#c8c8c8',
                    neutralSecondary: '#d0d0d0',
                    neutralPrimaryAlt: '#dadada',
                    neutralPrimary: '#ffffff',
                    neutralDark: '#f4f4f4',
                    black: '#f8f8f8',
                    white: '#323130',
                    hover: "#201f1e"
                };
                break;
            }
            case TeamsTheme.contrast: {
                theme.palette = {
                    themePrimary: '#00eaff',
                    themeLighterAlt: '#00090a',
                    themeLighter: '#002529',
                    themeLight: '#00464d',
                    themeTertiary: '#008c99',
                    themeSecondary: '#00cee0',
                    themeDarkAlt: '#19ecff',
                    themeDark: '#3defff',
                    themeDarker: '#70f3ff',
                    neutralLighterAlt: '#0b0b0b',
                    neutralLighter: '#151515',
                    neutralLight: '#252525',
                    neutralQuaternaryAlt: '#2f2f2f',
                    neutralQuaternary: '#373737',
                    neutralTertiaryAlt: '#595959',
                    neutralTertiary: '#4d4d00',
                    neutralSecondary: '#999900',
                    neutralPrimaryAlt: '#e0e000',
                    neutralPrimary: '#ffff00',
                    neutralDark: '#ffff3d',
                    black: '#ffff70',
                    white: '#000000',
                    hover: "#201f1e"
                };
                break;
            }
            default: {
                theme.palette = {
                    themePrimary: '#6362a5',
                    themeLighterAlt: '#f7f7fb',
                    themeLighter: '#e1e1f1',
                    themeLight: '#c9c8e4',
                    themeTertiary: '#9998c9',
                    themeSecondary: '#7271b0',
                    themeDarkAlt: '#595895',
                    themeDark: '#4b4a7e',
                    themeDarker: '#37375d',
                    neutralLighterAlt: '#faf9f8',
                    neutralLighter: '#f3f2f1',
                    neutralLight: '#edebe9',
                    neutralQuaternaryAlt: '#e1dfdd',
                    neutralQuaternary: '#d0d0d0',
                    neutralTertiaryAlt: '#c8c6c4',
                    neutralTertiary: '#a19f9d',
                    neutralSecondary: '#605e5c',
                    neutralPrimaryAlt: '#3b3a39',
                    neutralPrimary: '#323130',
                    neutralDark: '#201f1e',
                    black: '#000000',
                    white: '#ffffff',
                    hover: "#201f1e"
                };
                break;
            }
        }
        (window as any).__loadTheme(theme);
        ThemeManager.palette = theme.palette;
    }
}