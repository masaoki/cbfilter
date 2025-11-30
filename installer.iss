[Setup]
AppName=cbfilter
AppVersion=1.0.0
DefaultDirName={autopf}\cbfilter
DefaultGroupName=cbfilter
OutputDir=build
OutputBaseFilename=cbfilter-setup
Compression=lzma2
SolidCompression=yes
DisableDirPage=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\\Japanese.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\\ChineseSimplified.isl"
Name: "korean"; MessagesFile: "compiler:Languages\\Korean.isl"
Name: "vietnamese"; MessagesFile: "compiler:Languages\\Vietnamese.isl"
Name: "thai"; MessagesFile: "compiler:Languages\\Thai.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\\Spanish.isl"
Name: "german"; MessagesFile: "compiler:Languages\\German.isl"
Name: "french"; MessagesFile: "compiler:Languages\\French.isl"
Name: "italian"; MessagesFile: "compiler:Languages\\Italian.isl"
Name: "dutch"; MessagesFile: "compiler:Languages\\Dutch.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\\Portuguese.isl"
Name: "russian"; MessagesFile: "compiler:Languages\\Russian.isl"

[Files]
Source: "cbfilter.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "cbfilter.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "defconf_en.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: english
Source: "defconf_ja.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: japanese
Source: "defconf_zh.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: chinesesimplified
Source: "defconf_ko.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: korean
Source: "defconf_vi.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: vietnamese
Source: "defconf_th.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: thai
Source: "defconf_es.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: spanish
Source: "defconf_de.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: german
Source: "defconf_fr.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: french
Source: "defconf_it.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: italian
Source: "defconf_nl.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: dutch
Source: "defconf_pt.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: portuguese
Source: "defconf_ru.json"; DestDir: "{app}"; DestName: "defconf.json"; Flags: ignoreversion; Languages: russian
Source: "lang.ini"; DestDir: "{app}"; Flags: ignoreversion
Source: "apidef\\*.json"; DestDir: "{app}\\apidef"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\\cbfilter"; Filename: "{app}\\cbfilter.exe"; IconFilename: "{app}\\cbfilter.ico"
Name: "{commondesktop}\\cbfilter"; Filename: "{app}\\cbfilter.exe"; IconFilename: "{app}\\cbfilter.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional tasks:"; Flags: unchecked

[Run]
Filename: "{app}\\cbfilter.exe"; Description: "Launch cbfilter"; Flags: postinstall nowait skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{userappdata}\\cbfilter"
