export enum BrandIcons {
    Word = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/word_48x1.svg",
    PowerPoint = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/powerpoint_48x1.svg",
    Excel = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg",
    Pdf = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/pdf.svg",
    OneNote = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/onenote_48x1.svg",
    OneNotePage = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/onenote_48x1.svg",
    InfoPath = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/xsn.svg",
    Visio = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/vsdx.svg",
    Publisher = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/pub.svg",
    Project = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/mpp.svg",
    Access = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/accdb.svg",
    Mail = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/email.svg",
    Csv = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/xlsx.svg",
    Archive = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/zip.svg",
    Xps = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/genericfile.svg",
    Audio = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/audio.svg",
    Video = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/video.svg",
    Image = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/photo.svg",
    Text = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/txt.svg",
    Xml = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/xml.svg",
    Zip = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/zip.svg",
    Url = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/link.svg",
    Folder = "https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201207.001/assets/item-types/48/folder.svg",
}


export enum ApplicationType {
    Access,
    ASPX,
    Code,
    CSS,
    CSV,
    Excel,
    HTML,
    Image,
    Mail,
    OneNote,
    Pdf,
    PowerApps,
    PowerPoint,
    Project,
    Publisher,
    SASS,
    Visio,
    Word,
    Text,
    Video,
    Zip,
    Url
}
export interface IApplicationIcons {

    application: ApplicationType;
    extensions: string[];
    iconName: string;
    imageName: string[];
    cdnImageName?: string[];
}
export const ApplicationIconList: IApplicationIcons[] = [
    {
        application: ApplicationType.Access,
        extensions: ['accdb', 'accde', 'accdt', 'accdr', 'mdb'],
        iconName: 'AccessLogo',
        imageName: ['accdb'],
        cdnImageName: ["accdb"]
    },
    {
        application: ApplicationType.ASPX,
        extensions: ['aspx', 'master'],
        iconName: 'FileASPX',
        imageName: [],
        cdnImageName: ['spo']
    },
    {
        application: ApplicationType.Code,
        extensions: ['js', 'ts', 'cs'],
        iconName: 'FileCode',
        imageName: [],
        cdnImageName: ['code']
    },
    {
        application: ApplicationType.CSS,
        extensions: ['css'],
        iconName: 'FileCSS',
        imageName: [],
        cdnImageName: ['code']
    },
    {
        application: ApplicationType.CSV,
        extensions: ['csv'],
        iconName: 'ExcelDocument',
        imageName: ['csv'],
        cdnImageName: ['csv']
    },
    {
        application: ApplicationType.Excel,
        extensions: ['xls', 'xlt', 'xlm', 'xlsx', 'xlsm', 'xltx', 'xltm', 'ods'],
        iconName: 'ExcelDocument',
        imageName: ['xlsx', 'xls', 'xltx', 'ods'],
        cdnImageName: ['xlsx', 'xltx', 'ods']
    },
    {
        application: ApplicationType.HTML,
        extensions: ['html'],
        iconName: 'FileHTML',
        imageName: [],
        cdnImageName: ['html']
    },
    {
        application: ApplicationType.Image,
        extensions: ['jpg', 'jpeg', 'gif', 'png'],
        iconName: 'FileImage',
        imageName: [],
        cdnImageName: ['photo']
    },
    {
        application: ApplicationType.Mail,
        extensions: ['msg'],
        iconName: 'Mail',
        imageName: [],
        cdnImageName: ['email']
    },
    {
        application: ApplicationType.OneNote,
        extensions: ['one', 'onepkg', 'onetoc', 'onetoc2'],
        iconName: 'OneNote',
        imageName: ['one', 'onepkg', 'onetoc', 'onetoc2'],
        cdnImageName: ['one', 'onetoc', 'onetoc2']
    },
    {
        application: ApplicationType.Pdf,
        extensions: ['pdf', 'Pdf', 'PDF'],
        iconName: 'PDF',
        imageName: [],
        cdnImageName: ['pdf']
    },
    {
        application: ApplicationType.PowerApps,
        extensions: ['msapp'],
        iconName: 'PowerApps',
        imageName: []
    },
    {
        application: ApplicationType.PowerPoint,
        extensions: ['ppt', 'pot', 'pps', 'pptx', 'pptm', 'potx', 'potm', 'ppam', 'ppsx', 'ppsm', 'sldx', 'sldx'],
        iconName: 'PowerPointDocument',
        imageName: ['odp', 'potx', 'ppsx', 'pptx'],
        cdnImageName: ['pptx', 'odp', 'potx', 'ppsx']
    },
    {
        application: ApplicationType.Project,
        extensions: ['mpp', 'mpt', 'mpx', 'mpd'],
        iconName: 'ProjectLogoInverse',
        imageName: ['mpp', 'mpt'],
        cdnImageName: ['mpp', 'mpt']
    },
    {
        application: ApplicationType.Publisher,
        extensions: ['pub'],
        iconName: 'PublisherLogo',
        imageName: ['pub'],
        cdnImageName: ['pub']
    },
    {
        application: ApplicationType.SASS,
        extensions: ['scss', 'sass'],
        iconName: 'FileSass',
        imageName: [],
        cdnImageName: ['code']
    },
    {
        application: ApplicationType.Visio,
        extensions: ['vsd', 'vss', 'vst', 'vdx', 'vsx', 'vtx', 'vsdx'],
        iconName: 'VisioDocument',
        imageName: ['vsdx', 'vssx', 'vstx'],
        cdnImageName: ['vsdx', 'vssx', 'vstx']
    },
    {
        application: ApplicationType.Word,
        extensions: ['doc', 'dot', 'docx', 'docm', 'dotx', 'dotm', 'docb', 'odt'],
        iconName: 'WordDocument',
        imageName: ['docx', 'dotx', 'odt'],
        cdnImageName: ['docx', 'dotx', 'odt']
    },
    {
        application: ApplicationType.Text,
        extensions: ['txt'],
        iconName: 'Text',
        imageName: [],
        cdnImageName: ['text']
    },
    {
        application: ApplicationType.Video,
        extensions: ['mp4', 'avi', 'flv', 'wmv', 'mov'],
        iconName: 'Video',
        imageName: [],
        cdnImageName: ['Video']
    },
    {
        application: ApplicationType.Zip,
        extensions: ['zip'],
        iconName: 'Zip',
        imageName: [],
        cdnImageName: ['Zip']
    },
    {
        application: ApplicationType.Url,
        extensions: ['url'],
        iconName: 'Url',
        imageName: [],
        cdnImageName: ['Url']
    }
];