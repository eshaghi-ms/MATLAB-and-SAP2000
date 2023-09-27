% clean-up the workspace & command window
close all;
clear;
clc;
     
%% full path to the program executable
%% set it to the installation folder
disp('Set it to the installation folder')
ProgramPath = ...
    'C:\Program Files\Computers and Structures\SAP2000 19\sap2000.exe';

%% full path to API dll
%% set it to the installation folder
disp('Set it to the installation folder')
APIDLLPath = ...
    'C:\Program Files\Computers and Structures\SAP2000 19\sap2000v19.dll';

%% set it to the desired path of your model
disp('Set it to the desired path of your model')
ModelDirectory = 'C:\Code\SapData\Simple\GACSiAPIexample';

if ~exist(ModelDirectory, 'dir')
    mkdir(ModelDirectory)
end

ModelName = 'API_Simple_GA.sdb';
ModelPath = strcat(ModelDirectory, filesep, ModelName);

%% create OAPI helper object
disp('Create OAPI helper object')
a = NET.addAssembly(APIDLLPath);
helper = SAP2000v19.Helper;
helper = NET.explicitCast(helper,'SAP2000v19.cHelper');

%% create Sap2000 object
disp('Create Sap2000 object')
SapObject = helper.CreateObject(ProgramPath);
SapObject = NET.explicitCast(SapObject,'SAP2000v19.cOAPI');
helper = 0;

%% Start Sap2000 application
disp('Start Sap2000 application')
Ret = SapObject.ApplicationStart;
if Ret ~= 0
    disp('! Error at Start Sap2000 application')
end

%% create SapModel object
disp('Create SapModel object')
SapModel = NET.explicitCast(SapObject.SapModel,'SAP2000v19.cSapModel');

%% initialize model
disp('Initialize model')
Ret = SapModel.InitializeNewModel(SAP2000v19.eUnits.kgf_m_C);
if Ret ~= 0
    disp('! Error at Initialize model')
end

%% create new blank model
disp('Create new blank model')
File = NET.explicitCast(SapModel.File,'SAP2000v19.cFile');
Ret = File.NewBlank;
if Ret ~= 0
    disp('! Error at Create new blank model')
end

%% define material property
disp('Define material property')
PropMaterial = ...
    NET.explicitCast(SapModel.PropMaterial,'SAP2000v19.cPropMaterial');

Ret = PropMaterial.SetMaterial('ST37Roller', SAP2000v19.eMatType.Steel);
if Ret ~= 0
    disp('! Error at Define material property ST37Roller')
end
Ret = PropMaterial.SetMaterial('C21', SAP2000v19.eMatType.Concrete);
if Ret ~= 0
    disp('! Error at Define material property C21')
end
Ret = PropMaterial.SetMaterial('C0' , SAP2000v19.eMatType.Concrete);
if Ret ~= 0
    disp('! Error at Define material property C0')
end

%% assign isotropic mechanical properties to material
disp('Assign isotropic mechanical properties to material')
Ret = PropMaterial.SetMPIsotropic('ST37Roller', 2E+10, 0.3, 1.17E-05);
if Ret ~= 0
    disp('! Error at SetMPIsotropic ST37Roller')
end
Ret = PropMaterial.SetMPIsotropic('C21' , 2.495+9, 0.15, 9.9E-06);
if Ret ~= 0
    disp('! Error at SetMPIsotropic C21')
end
Ret = PropMaterial.SetMPIsotropic('C0'  , 2.495+9, 0.15, 9.9E-06);
if Ret ~= 0
    disp('! Error at SetMPIsotropic C0')
end

Ret = PropMaterial.SetOSteel_1('ST37Roller', 24E+6,...
                37E+6, 28.8E+6, 44.40E+6, 1, 1, 0.015, 0.11, 0.17, -0.1);
if Ret ~= 0
    disp('! Error at SetOSteel_1 ST37Roller')
end
Ret = PropMaterial.SetOConcrete_1('C21', 21E+5, ...
                false, 0, 2, 2, 2.219E-03, 5.000E-03, -0.1, 0, 0);
if Ret ~= 0
    disp('! Error at SetOConcrete_1 C21')
end
Ret = PropMaterial.SetOConcrete_1('C0' , 21E+5, ...
                false, 0, 2, 2, 2.219E-03, 5.000E-03, -0.1, 0, 0);              
if Ret ~= 0
    disp('! Error at SetOConcrete_1 C0')
end
SteelWPerUnitV = 7850;
Ret = PropMaterial.SetWeightAndMass('ST37Roller', 1, 7850);
if Ret ~= 0
    disp('! Error at SetWeightAndMass ST37Roller')
end
Ret = PropMaterial.SetWeightAndMass('C21', 1, 2500);
if Ret ~= 0
    disp('! Error at SetWeightAndMass C21')
end
Ret = PropMaterial.SetWeightAndMass('C0' , 1, 0   );
if Ret ~= 0
    disp('! Error at SetWeightAndMass C0')
end

Ret = PropMaterial.Delete('4000psi');
if Ret ~= 0
    disp('! Error at PropMaterial.Delete 4000psi')
end
Ret = PropMaterial.Delete('A992Fy50');
if Ret ~= 0
    disp('! Error at PropMaterial.Delete A992Fy50')
end

%% define frame section property
disp('Define frame section property')

[~, W] = xlsread('Data.xlsx');
Path = 'D:\Program Files\Computers and Structures\SAP2000 19\euro.pro';
PropFrame = NET.explicitCast(SapModel.PropFrame,'SAP2000v19.cPropFrame');

Counter = numel(W);

WSection(Counter).Name      = W{Counter};
WSection(Counter).Area      = 0;
WSection(Counter).As2       = 0;
WSection(Counter).As3       = 0;
WSection(Counter).Torsion   = 0;
WSection(Counter).I22       = 0;
WSection(Counter).I33       = 0;
WSection(Counter).S22       = 0;
WSection(Counter).S33       = 0;
WSection(Counter).Z22       = 0;
WSection(Counter).Z33       = 0;
WSection(Counter).R22       = 0;
WSection(Counter).R33       = 0;
WSection(Counter).WPerUnitV = 0;

for i=1:Counter
    
    WSection(i).Name      = W{i};
    WSection(i).Area      = 0;
    WSection(i).As2       = 0;
    WSection(i).As3       = 0;
    WSection(i).Torsion   = 0;
    WSection(i).I22       = 0;
    WSection(i).I33       = 0;
    WSection(i).S22       = 0;
    WSection(i).S33       = 0;
    WSection(i).Z22       = 0;
    WSection(i).Z33       = 0;
    WSection(i).R22       = 0;
    WSection(i).R33       = 0;
    WSection(i).WPerUnitV = SteelWPerUnitV;
    
    PropFrame.ImportProp(WSection(i).Name, 'ST37Roller', Path, ...
                                                        WSection(i).Name);
    
    [Ret, WSection(i).Area    , WSection(i).As2 , ...
          WSection(i).As3     , WSection(i).Torsion , WSection(i).I22 , ...
          WSection(i).I33     , WSection(i).S22     , WSection(i).S33 , ...
          WSection(i).Z22     , WSection(i).Z33     , WSection(i).R22 , ...
          WSection(i).R33]...
          = PropFrame.GetSectProps( ...
          WSection(i).Name    , WSection(i).Area    , WSection(i).As2 , ...
          WSection(i).As3     , WSection(i).Torsion , WSection(i).I22 , ...
          WSection(i).I33     , WSection(i).S22     , WSection(i).S33 , ...
          WSection(i).Z22     , WSection(i).Z33     , WSection(i).R22 , ...
          WSection(i).R33);
      if Ret == 0
          disp([WSection(i).Name '   OK'])
      else 
          disp([WSection(i).Name '   NOT OK!'])
      end
end

%% define deck section property
disp('define deck section property')
PropArea = NET.explicitCast(SapModel.PropArea,'SAP2000v19.cPropArea');
Thickness = 0.15;
Ret = PropArea.SetShell_1('SHELL1', 1, true, 'C21', 0, Thickness, Thickness);
if Ret ~= 0
    disp('! Error at PropArea.SetShell_1 SHELL1')
end
Ret = PropArea.SetShell_1('SHELL0', 1, true, 'C0' , 0, 0.15, 0.15);
if Ret ~= 0
    disp('! Error at PropArea.SetShell_1 SHELL0')
end

%% add load patterns
disp('add load patterns')
LoadPatterns = ...
    NET.explicitCast(SapModel.LoadPatterns,'SAP2000v19.cLoadPatterns');

Ret = LoadPatterns.Add('LiveRR'  , SAP2000v19.eLoadPatternType.ReduceLive);
if Ret ~= 0
    disp('! Error at LoadPatterns.Add LiveRR')
end
Ret = LoadPatterns.Add('LivenR'  , SAP2000v19.eLoadPatternType.Live      );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add LivenR')
end
Ret = LoadPatterns.Add('Lr'      , SAP2000v19.eLoadPatternType.Rooflive  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add Lr')
end
Ret = LoadPatterns.Add('Leq'     , SAP2000v19.eLoadPatternType.Live      );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add Leq')
end
Ret = LoadPatterns.Add('MASS'    , SAP2000v19.eLoadPatternType.Other     );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add MASS')
end
Ret = LoadPatterns.Add('EX'      , SAP2000v19.eLoadPatternType.Quake     );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add EX')
end
Ret = LoadPatterns.Add('EY'      , SAP2000v19.eLoadPatternType.Quake     );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add EY')
end
Ret = LoadPatterns.Add('NDeadX'  , SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NDeadX')
end
Ret = LoadPatterns.Add('NDeadY'  , SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NDeadY')
end
Ret = LoadPatterns.Add('NLiveRRX', SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLiveRRX')
end
Ret = LoadPatterns.Add('NLiveRRY', SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLiveRRY')
end
Ret = LoadPatterns.Add('NLivenRX', SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLivenRX')
end
Ret = LoadPatterns.Add('NLivenRY', SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLivenRY')
end
Ret = LoadPatterns.Add('NLrX'    , SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLrX')
end
Ret = LoadPatterns.Add('NLrY'    , SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLrY')
end
Ret = LoadPatterns.Add('NLeqX'   , SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLeqX')
end
Ret = LoadPatterns.Add('NLeqY'   , SAP2000v19.eLoadPatternType.Notional  );
if Ret ~= 0
    disp('! Error at LoadPatterns.Add NLeqY')
end

%% modification of load patterns
input('Set Notional load setting and input a number  ');
disp('modification of load patterns')
AutoSeismic = ...
    NET.explicitCast(LoadPatterns.AutoSeismic,'SAP2000v19.cAutoSeismic');
Function = NET.explicitCast(SapModel.Func,'SAP2000v19.cFunction');
FuncRS   = NET.explicitCast(Function.FuncRS,'SAP2000v19.cFunctionRS');

Ret = AutoSeismic.SetUserCoefficient('EX', 1, 0, false, 0, 0, 0.126, 1);
if Ret ~= 0
    disp('! Error at AutoSeismic.SetUserCoefficient EX')
end
Ret = AutoSeismic.SetUserCoefficient('EY', 2, 0, false, 0, 0, 0.126, 1);
if Ret ~= 0
    disp('! Error at AutoSeismic.SetUserCoefficient EY')
end

Ret = Function.Delete('UNIFRS');
if Ret ~= 0
    disp('! Error at Function.Delete UNIFRS')
end

B1 = NET.createArray('System.Double',100);
T  = NET.createArray('System.Double',100);
B  = NET.createArray('System.Double',100);
NAB  = NET.createArray('System.Double',100);
T(1) = 0.05;
for NumberOfComb = 1:100
    if T(NumberOfComb) <= 0.15
        B1(NumberOfComb) = 1.1+1.65*T(NumberOfComb)/0.15;
    else if ((T(NumberOfComb) > 0.15) && (T(NumberOfComb) <= 0.7))
        B1(NumberOfComb) = 2.75;
        else 
            B1(NumberOfComb) = 2.75*0.7/T(NumberOfComb);
        end
    end
    
    if T(NumberOfComb) <= 0.7
        NAB(NumberOfComb) = 1.0;
    else if ((T(NumberOfComb) > 0.7) && (T(NumberOfComb) <= 4.0))
        NAB(NumberOfComb) = (0.7/3.3)*(T(NumberOfComb)-0.7)+1.0;
        else 
            NAB(NumberOfComb) = 1.70;
        end
    end
    B(NumberOfComb) = B1(NumberOfComb)*NAB(NumberOfComb);
    if NumberOfComb ~= 100
        T(NumberOfComb+1) = T(NumberOfComb) + 0.05;
    end
end

Ret = FuncRS.SetUser('Func1', 100, T, B, 0.05);
if Ret ~= 0
    disp('! Error at FuncRS.SetUser Func1')
end

%% Add load Cases
disp('Add load Cases')
LoadCases = NET.explicitCast(SapModel.LoadCases,'SAP2000v19.cLoadCases');
ResponseSpectrum = NET.explicitCast...
         (LoadCases.ResponseSpectrum,'SAP2000v19.cCaseResponseSpectrum');

ExistingLoadCase = ['SX '; 'SPX'; 'SY '; 'SPY'];

for i = 1:4
    Ret = ResponseSpectrum.SetCase(ExistingLoadCase(i,:));
    if Ret ~= 0
        disp(['! Error at ResponseSpectrum.SetCase i = ' num2str(i)])
    end
    Ret = ResponseSpectrum.SetModalComb_1(ExistingLoadCase(i,:), 1, 1, 0, 1);
    if Ret ~= 0
        disp(['! Error at ResponseSpectrum.SetModalComb_1 i = ' num2str(i)])
    end
    Ret = ResponseSpectrum.SetDirComb(ExistingLoadCase(i,:), 1, 0);
    if Ret ~= 0
        disp(['! Error at ResponseSpectrum.SetDirComb i = ' num2str(i)])
    end
    Ret = ResponseSpectrum.SetEccentricity(ExistingLoadCase(i,:), 0);
    if Ret ~= 0
        disp(['! Error at ResponseSpectrum.SetEccentricity i = ' num2str(i)])
    end
end
                                                           
U    = NET.createArray('System.String',1);
Func = NET.createArray('System.String',1);
SF   = NET.createArray('System.Double',1);
CSys = NET.createArray('System.String',1);
Ang  = NET.createArray('System.Double',1);

U   (1) = 'U1';
Func(1) = 'Func1';
SF  (1) = 0.20601;
CSys(1) = 'Global';
Ang (1) = 0;

Ret = ResponseSpectrum.SetLoads('SX' , 1, U, Func, SF, CSys, Ang);
if Ret ~= 0
    disp('! Error at ResponseSpectrum.SetLoads SX')
end
SF  (1) = 0.6867;
Ret = ResponseSpectrum.SetLoads('SPX', 1, U, Func, SF, CSys, Ang);
if Ret ~= 0
    disp('! Error at ResponseSpectrum.SetLoads SPX')
end
U   (1) = 'U2';
SF  (1) = 0.20601;
Ret = ResponseSpectrum.SetLoads('SY' , 1, U, Func, SF, CSys, Ang);
if Ret ~= 0
    disp('! Error at ResponseSpectrum.SetLoads SY')
end
SF  (1) = 0.6867;
Ret = ResponseSpectrum.SetLoads('SPY', 1, U, Func, SF, CSys, Ang);
if Ret ~= 0
    disp('! Error at ResponseSpectrum.SetLoads SPY')
end

%% add load combination
disp('Add load combination')
LoadCombinationFactor = xlsread('LoadCombinationFactor.xlsx');
LoadName = ['DEAD    '; 'LiveRR  '; 'LivenR  '; 'Lr      '; 'Leq     '; ...
            'EX      '; 'EY      '; 'NdeadX  '; 'NdeadY  '; 'NLiveRRX'; ...
            'NLiveRRY'; 'NLivenRX'; 'NLivenRY'; 'NLrX    '; 'NlrY    '; ...
            'NLeqX   '; 'NLeqY   '; 'SX      ';	'SY      '; 'SPX     '; ...
            'SPY     '];

RespCombo = NET.explicitCast(SapModel.RespCombo,'SAP2000v19.cCombo');
                
for NumberOfComb = 1:22
    CombName = ['Comb' num2str(NumberOfComb)];
    Ret = RespCombo.Add(CombName, 0);
    if Ret ~= 0
        disp('! Error at Add Load Combination')        
    end
    for NumberOfLoad = 1:21
        if LoadCombinationFactor(NumberOfComb, NumberOfLoad) ~= 0
        Ret = RespCombo.SetCaseList(CombName, ...
            SAP2000v19.eCNameType.LoadCase, LoadName(NumberOfLoad,:), ...
            LoadCombinationFactor(NumberOfComb, NumberOfLoad));
            if Ret ~= 0
                disp('! Error at SetCaseList in Load Combination')
            end
        end
    end
end

%% Define mass source
disp('Define mass source')
SourceMass = ...
    NET.explicitCast(SapModel.SourceMass,'SAP2000v19.cMassSource');

LoadPat = NET.createArray('System.String',6);
LoadPat(1) = 'DEAD';
LoadPat(2) = 'LiveRR';
LoadPat(3) = 'LivenR';
LoadPat(4) = 'Lr';
LoadPat(5) = 'Leq';
LoadPat(6) = 'MASS';

SF = NET.createArray('System.Double',6);
SF(1) = 1.0;
SF(2) = 0.2;
SF(3) = 0.2;
SF(4) = 0.2;
SF(5) = 1.0;
SF(6) = 1.0;

Ret = SourceMass.SetMassSource...
    ('MSSSRC1', false, false, true, false, 6, LoadPat, SF);
if Ret ~= 0
    disp('! SourceMass.SetMassSource MSSSRC1')
end

%% Add beam object by coordinates
disp('Add beam object by coordinates')
FrameObj = NET.explicitCast(SapModel.FrameObj,'SAP2000v19.cFrameObj');
View = NET.explicitCast(SapModel.View,'SAP2000v19.cView');

NumberOfStory = 9;
PointData  = xlsread('Point.xlsx');
NumOfPoint = size(PointData,1);
ElementData  = xlsread('Element.xlsx');
NumOfElement = size(ElementData,1);
NumOfAllBeam = NumberOfStory*NumOfElement;
NAB = NumOfAllBeam;
NE = NumOfElement;


Element(NAB).Type    = 'BEAM';
Element(NAB).I       = 0;
Element(NAB).J       = 0;
Element(NAB).XI      = 0;
Element(NAB).XJ      = 0;
Element(NAB).YI      = 0;
Element(NAB).YJ      = 0;
Element(NAB).ZI      = 0;
Element(NAB).ZJ      = 0;
Element(NAB).Length  = 0;
Element(NAB).Name    = System.String(' ');
Element(NAB).Section = WSection(1);
RnadMax = size(WSection,2);
Element(NumberOfStory*NE).Volume  = 0;

FrameName{NAB} = System.String(' ');

for i = 1:NumberOfStory
    for j = 1:NE
        
        R = randi(RnadMax);
        
        Element((i-1)*NE+j).Type = 'BEAM';
        Element((i-1)*NE+j).I    = ElementData(j,1);
        Element((i-1)*NE+j).J    = ElementData(j,2);
        Element((i-1)*NE+j).XI   = PointData(ElementData(j,1),1);
        Element((i-1)*NE+j).XJ   = PointData(ElementData(j,2),1);
        Element((i-1)*NE+j).YI   = PointData(ElementData(j,1),2);
        Element((i-1)*NE+j).YJ   = PointData(ElementData(j,2),2);
        Element((i-1)*NE+j).ZI   = 3*i;
        Element((i-1)*NE+j).ZJ   = 3*i;
        Element((i-1)*NE+j).Length = ...
            sqrt((Element((i-1)*NE+j).XJ - Element((i-1)*NE+j).XI)^2 + ...
            (Element((i-1)*NE+j).YJ - Element((i-1)*NE+j).YI)^2 + ...
            (Element((i-1)*NE+j).ZJ - Element((i-1)*NE+j).ZI)^2); 
        Element((i-1)*NE+j).Name   = System.String(...
            [Element((i-1)*NE+j).Type num2str((i-1)*NE+j)]);
        Element((i-1)*NE+j).Section = WSection(R);
                        
        Ret = FrameObj.AddByCoord(...
            Element((i-1)*NE+j).XI,   Element((i-1)*NE+j).YI,  ...
            Element((i-1)*NE+j).ZI,   Element((i-1)*NE+j).XJ,  ...
            Element((i-1)*NE+j).YJ,   Element((i-1)*NE+j).ZJ,  ...
            Element((i-1)*NE+j).Name, Element((i-1)*NE+j).Section.Name, ...
            Element((i-1)*NE+j).Name, 'Global');
        if Ret ~= 0
            disp(['! Add Beam; NumberOfStory = ' num2str(i) '  Element = ' num2str(j)])
        end
        
        Element((i-1)*NE+j).Volume = ...
            Element((i-1)*NE+j).Section.Area*Element((i-1)*NE+j).Length;
     
    end
end
View.RefreshView();
if Ret ~= 0
    disp('!Add beam object View.RefreshView')
end

%% Add column object by coordinates
disp('Add column object by coordinates')

Z = NumberOfStory*NumOfElement;
NumOfColumn = NumOfPoint;
NumOfAllColumn = NumberOfStory*NumOfColumn;
NAC = NumOfAllColumn;
NC = NumOfColumn;

Element(Z + NAC).Type   = 'COLUMN';
Element(Z + NAC).I      = 0;
Element(Z + NAC).J      = 0;
Element(Z + NAC).XI     = 0;
Element(Z + NAC).XJ     = 0;
Element(Z + NAC).YI     = 0;
Element(Z + NAC).YJ     = 0;
Element(Z + NAC).ZI     = 0;
Element(Z + NAC).ZJ     = 0;
Element(Z + NAC).Length = 0;
Element(Z + NAC).Name   = System.String(' ');
Element(NumberOfStory*NE).Section = WSection(1);
Element(NumberOfStory*NE).Volume  = 0;


for i = 1:NumberOfStory
    for j = 1:NC
        
        R = randi(RnadMax);
        
        Element(Z + (i-1)*NC+j).Type = 'COLUMN';
        Element(Z + (i-1)*NC+j).I    = 1;
        Element(Z + (i-1)*NC+j).J    = NC+1;
        Element(Z + (i-1)*NC+j).XI   = PointData(j,1);
        Element(Z + (i-1)*NC+j).XJ   = PointData(j,1);
        Element(Z + (i-1)*NC+j).YI   = PointData(j,2);
        Element(Z + (i-1)*NC+j).YJ   = PointData(j,2);
        Element(Z + (i-1)*NC+j).ZI   = 3*(i-1);
        Element(Z + (i-1)*NC+j).ZJ   = 3*i;
        Element(Z + (i-1)*NC+j).Length = sqrt(...
            (Element(Z + (i-1)*NC+j).XJ - Element(Z + (i-1)*NC+j).XI)^2 + ...
            (Element(Z + (i-1)*NC+j).YJ - Element(Z + (i-1)*NC+j).YI)^2 + ...
            (Element(Z + (i-1)*NC+j).ZJ - Element(Z + (i-1)*NC+j).ZI)^2);
                         
        Element(Z + (i-1)*NC+j).Name   = System.String(...
            [Element(Z + (i-1)*NC+j).Type num2str((i-1)*NC+j)]);
        
        Element(Z + (i-1)*NC+j).Section = WSection(R);
        
        
        Ret = FrameObj.AddByCoord( ...
            Element(Z + (i-1)*NC+j).XI  , Element(Z + (i-1)*NC+j).YI,   ...
            Element(Z + (i-1)*NC+j).ZI  , Element(Z + (i-1)*NC+j).XJ,   ...
            Element(Z + (i-1)*NC+j).YJ  , Element(Z + (i-1)*NC+j).ZJ,   ...
            Element(Z + (i-1)*NC+j).Name, Element(Z + (i-1)*NC+j).Section.Name,...
            Element(Z + (i-1)*NC+j).Name, 'Global');
        
        if Ret ~= 0
            disp(['! Add Column; NumberOfStory = ' num2str(i) '  Element = ' num2str(j)])
        end
                         
        Element(Z + (i-1)*NC+j).Volume = ...
            Element(Z + (i-1)*NC+j).Section.Area*Element(Z + (i-1)*NC+j).Length;
                         
    end 
end
 View.RefreshView();

 %% Add area object by coordinates
disp('Add area object by coordinates')

AreaObj    = NET.explicitCast(SapModel.AreaObj,'SAP2000v19.cAreaObj');
AreaData   = xlsread('Area.xlsx');
NumOfArea  = size(AreaData,1);
NA = NumOfArea;


Area(NumberOfStory*NA).Type   = 'AREA';
Area(NumberOfStory*NA).PointNum     = 0;
Area(NumberOfStory*NA).X      = zeros(1,3);
Area(NumberOfStory*NA).Y      = zeros(1,3);
Area(NumberOfStory*NA).Z      = zeros(1,3);
Area(NumberOfStory*NA).Name   = System.String(' ');

AreaName{NumberOfStory*NA} = System.String('');
AreaPoint = AreaData(1,1);

 for i = 1:NumberOfStory
     for j = 1:NA
        Area((i-1)*NA+j).PointNum = AreaData(j,1);
        Area((i-1)*NA+j).Name     = System.String(...
            [Area(NumberOfStory*NA).Type num2str((i-1)*NA+j)]);        
        Area((i-1)*NA+j).X = zeros(1,Area((i-1)*NA+j).PointNum);
        Area((i-1)*NA+j).Y = zeros(1,Area((i-1)*NA+j).PointNum);
        Area((i-1)*NA+j).Z = zeros(1,Area((i-1)*NA+j).PointNum);
        for m = 1:Area((i-1)*NA+j).PointNum
            Area((i-1)*NA+j).X(m) = PointData(AreaData(j,m+1),1);
            Area((i-1)*NA+j).Y(m) = PointData(AreaData(j,m+1),2);
            Area((i-1)*NA+j).Z(m) = 3*i;
        end
        
        Ret = AreaObj.AddByCoord(Area((i-1)*NA+j).PointNum, ...
            Area((i-1)*NA+j).X, Area((i-1)*NA+j).Y, Area((i-1)*NA+j).Z, ...
            Area((i-1)*NA+j).Name, 'SHELL1', Area((i-1)*NA+j).Name);
        if Ret ~= 0
            disp(['! Add Area; NumberOfStory = ' num2str(i) '  Element = ' num2str(j)])
        end
        
     end
 end
       
%% Set Diaphragm
disp('Set Diaphragm')

ConstraintDef = ...
    NET.explicitCast(SapModel.ConstraintDef,'SAP2000v19.cConstraint');
PointObj = NET.explicitCast(SapModel.PointObj,'SAP2000v19.cPointObj');
DesignSteel = NET.explicitCast(SapModel.DesignSteel,'SAP2000v19.cDesignSteel');
AISC360_10 = NET.explicitCast(DesignSteel.AISC360_10, 'SAP2000v19.cDStAISC360_10');

Ret = ConstraintDef.SetDiaphragm('Diaphragm1', SAP2000v19.eConstraintAxis.Z);
if Ret ~= 0
	disp('! Error at SetDiaphragm')
end

NumOfAllPoint = NumberOfStory*NumOfPoint;
PointName = NET.createArray('System.String',NumOfAllPoint);
[Ret, NumOfAllPoint, PointName] = PointObj.GetNameList(1, PointName);
if Ret ~= 0
	disp('! Error at PointObj.GetNameList')
end

SelectObj = NET.explicitCast(SapModel.SelectObj,'SAP2000v19.cSelect');

for i=1:NumberOfStory+1
    Ret = SelectObj.PlaneXY(PointName(NumOfPoint*(i-1)+1));
    if Ret ~= 0
        disp(['! Error at SelectObj.PlaneXY NumberOfStory = ' num2str(i)])
    end
    Ret = PointObj.SetConstraint(...
        '', 'Diaphragm1', SAP2000v19.eItemType.SelectedObjects);
    if Ret ~= 0
        disp(['! Error at PointObj.SetConstraint NumberOfStory = ' num2str(i)])
    end
    Ret = AISC360_10.SetOverwrite('', 19, 0.1, SAP2000v19.eItemType.SelectedObjects);
    if Ret ~= 0
        disp(['! Error at AISC360_10.Overwrite NumberOfStory = ' num2str(i)])
    end
    Ret = SelectObj.ClearSelection;
    if Ret ~= 0
        disp(['! Error at SelectObj.ClearSelection NumberOfStory = ' num2str(i)])
    end
end

Ret = SelectObj.PlaneXY(PointName(1));
if Ret ~= 0
    disp('! Error at SelectObj.PlaneXY NumberOfStory = 1')
end
Ret = PointObj.SetConstraint(...
        '', 'Diaphragm1', SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
    disp('! Error at PointObj.SetConstraint NumberOfStory = 1')
end
Ret = SelectObj.ClearSelection;
if Ret ~= 0
    disp('! Error at SelectObj.ClearSelection NumberOfStory = 1')
end

%% SetRestraint
disp('Set Restraint')

Ret = SelectObj.PlaneXY(PointName(NumOfPoint*NumberOfStory+1));
if Ret ~= 0
    disp('! Error at SelectObj.PlaneXY ')
end
Restraint = NET.createArray('System.Boolean',6);

for i = 1 : 6
    Restraint(i) = true();
end

Ret = PointObj.SetRestraint('', Restraint, SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
    disp('! Error at PointObj.SetRestraint')
end

Ret = SelectObj.ClearSelection;
if Ret ~= 0
    disp(['! Error at SelectObj.ClearSelection NumberOfStory = ' num2str(i)])
end

%% Add Wall Load
disp('Add Wall Loading')

AreaOnlyPoint = AreaData;
AreaOnlyPoint(:,1) = [];
TypeOfElement(NE) = 0;

for i = 1:NE
    TypeOfElement(i) = sum(sum(ismember(AreaOnlyPoint, ElementData(i,:))));
    if TypeOfElement(i) < 5
        for j = 1:NumberOfStory-1
            Ret = FrameObj.SetLoadDistributed(Element((j-1)*NE+i).Name,...
                'DEAD', 1, 10, 0, 1, 530, 530);
            if Ret ~= 0
                disp('! Error at FrameObj.SetLoadDistributed DEAD')
                disp([' Number Of Story   = ' num2str(j)])
                disp([' Number of Element = ' num2str(i)])
            end
        end
        
        Ret = FrameObj.SetLoadDistributed(Element((NumberOfStory-1)*NE+i).Name,...
                'DEAD', 1, 10, 0, 1, 250, 250);
        if Ret ~= 0
            disp('! Error at FrameObj.SetLoadDistributed')
            disp(' Number Of Story   = 12')
            disp([' Number of Element = ' num2str(i)])
        end
        
        Ret = FrameObj.SetLoadDistributed(Element((NumberOfStory-1)*NE+i).Name,...
                'MASS', 1, 10, 0, 1, 265, 265);
        if Ret ~= 0
            disp('! Error at FrameObj.SetLoadDistributed')
            disp(' Number Of Story   = 12')
            disp([' Number of Element = ' num2str(i)])
        end    
    end
end

%% Area Loading
disp('Area Loading')

Ret = SelectObj.All;
if Ret ~= 0
	disp('! Error at SelectObj.All')
end    
Ret = SelectObj.PlaneXY(PointName(NumOfPoint*(NumberOfStory-1)+1), true());
if Ret ~= 0
	disp('! Error at SelectObj.PlaneXY Story')
end    
Ret = AreaObj.SetLoadUniform('', 'DEAD', 300, 10, true(), 'Global',...
                                SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
	disp('! Error at AreaObj.SetLoadUniform DEAD')
end    
Ret = AreaObj.SetLoadUniform('', 'Leq' , 115, 10, true(), 'Global',...
                                SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
	disp('! Error at AreaObj.SetLoadUniform Leq')
end    
Ret = AreaObj.SetLoadUniform('', 'LiveRR', 200, 10, true(), 'Global',...
                                SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
	disp('! Error at AreaObj.SetLoadUniform LiveRR')
end    

Ret = SelectObj.ClearSelection;
if Ret ~= 0
	disp('! Error at SelectObj.ClearSelection')
end    

Ret = SelectObj.PlaneXY(PointName(NumOfPoint*(NumberOfStory-1)+1));
if Ret ~= 0
	disp('! Error at SelectObj.PlaneXY Roof')
end    
Ret = AreaObj.SetLoadUniform('', 'DEAD', 345, 10, true(), 'Global',...
                                SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
	disp('! Error at AreaObj.SetLoadUniform Roof DEAD')
end    
Ret = AreaObj.SetLoadUniform('', 'Lr' , 150, 10, true(), 'Global',...
                                SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
	disp('! Error at AreaObj.SetLoadUniform Roof Lr')
end    
Ret = AreaObj.SetLoadUniform('', 'MASS', 57.5, 10, true(), 'Global',...
                                SAP2000v19.eItemType.SelectedObjects);
if Ret ~= 0
	disp('! Error at AreaObj.SetLoadUniform Roof MASS')
end    
                            
Ret = SelectObj.ClearSelection;
if Ret ~= 0
	disp('! Error at SelectObj.ClearSelection')
end    

%% Divide Column
disp('Divide Column')

for i = 1:NumberOfStory
    BaseHight = 3*(i-1);
    for j = 1:NumOfPoint
        X = PointData(j,1);
        Y = PointData(j,2);
        for m = 1:4
            Name = ['Point' num2str(NumOfAllPoint+...
                                                NumOfPoint*(i-1)+j*3+m)];
            Hight = BaseHight + 0.5*m ;
            Ret = PointObj.AddCartesian(X, Y, Hight, Name, Name);
            if Ret ~= 0
                disp(['! Error at Divide Column NumberOfStory = ' num2str(i) 'NumOfPoint = ' num2str(j)])
            end 
        end   
    end
end


%% P-Delta
disp('P-Delta')

StaticNonlinear = NET.explicitCast...
         (LoadCases.StaticNonlinear,'SAP2000v19.cCaseStaticNonlinear');

LoadType = NET.createArray('System.String',5);
LoadName = NET.createArray('System.String',5);
PDeltaSF = [1.2, 0.5, 1.0, 1.0, 0.2];
for i = 1:5
    LoadType(i) = 'Load';
end
LoadName(1) = 'DEAD';
LoadName(2) = 'LiveRR';
LoadName(3) = 'LivenR';
LoadName(4) = 'Leq';
LoadName(5) = 'Lr';

Ret = StaticNonlinear.SetCase('P-Delta');
if Ret ~= 0
    disp('! Error at StaticNonlinear.SetCase')
end
Ret = StaticNonlinear.SetGeometricNonlinearity('P-Delta', 1);
if Ret ~= 0
    disp('! Error at StaticNonlinear.SetGeometricNonlinearity')
end
Ret = StaticNonlinear.SetLoads('P-Delta', 5, LoadType, LoadName, PDeltaSF); 
if Ret ~= 0
    disp('! Error at StaticNonlinear.SetLoads')
end

%% Dynamic Modal
disp('Dynamic Modal')

ModalEigen = NET.explicitCast...
         (LoadCases.ModalEigen,'SAP2000v19.cCaseModalEigen');

Ret = ModalEigen.SetCase('Modal');
if Ret ~= 0
    disp('! Error at ModalEigen.SetCase')
end
Ret = ModalEigen.SetInitialCase('Modal', 'P-Delta');
if Ret ~= 0
    disp('! Error at ModalEigen.SetInitialCase')
end
Ret = ModalEigen.SetNumberModes('Modal', 33, 3);
if Ret ~= 0
    disp('! Error at ModalEigen.SetNumberModes')
end

%% Save model
disp('Save model')

Ret = File.Save(ModelPath);
if Ret ~= 0
    disp('! Error at Save model')
end

Ret = SapObject.Hide;
if Ret ~= 0
    disp('! Error at SapObject.Hide')
end

%% Define Analyze & Design Object & Array
disp('Define Analyze & Design Object & Array')

Analyze = NET.explicitCast(SapModel.Analyze,'SAP2000v19.cAnalyze');

DesignedNumberItems = numel(Element);
DesignedFrameName = NET.createArray('System.String',DesignedNumberItems);
DesignedRatio = NET.createArray('System.Double',DesignedNumberItems);
DesignedRatioType = NET.createArray('System.Int32',DesignedNumberItems);
DesignedLocation = NET.createArray('System.Double',DesignedNumberItems);
DesignedComboName = NET.createArray('System.String',DesignedNumberItems);
DesignedErrorSummary = NET.createArray('System.String',DesignedNumberItems);
DesignedWarningSummary = NET.createArray('System.String',DesignedNumberItems);

Ret = AISC360_10.SetPreference(1, 2);
if Ret ~= 0
    disp('! Error at Analyze.SetSolverOption')
end

%Ret = Analyze.SetSolverOption_1(2,2,false());
%if Ret ~= 0
%    disp('! Error at Analyze.SetSolverOption')
%end
%% Define DataBase

empty_individual.DeckThickness = Thickness;
empty_individual.Seismic.A  = 0.35;
empty_individual.Seismic.I  = 1;
empty_individual.Seismic.T  = 1.175;
empty_individual.Seismic.T0 = 0.15;
empty_individual.Seismic.S  = 1.75;
empty_individual.Geometry.PointData = PointData;
empty_individual.Geometry.ElementData = ElementData;
empty_individual.Geometry.NumberOfStory = NumberOfStory;
empty_individual.Geometry.Hight = 3;
empty_individual.ElementSectionNumber = [];
empty_individual.BeamSectionNumber = [];
empty_individual.ColumnSectionNumber = [];
empty_individual.Load.Distributed.Roof = 530;
empty_individual.Load.Distributed.Story = 250;
empty_individual.Load.Distributed.Mass = 265;
empty_individual.Load.Area.Story.Dead = 300;
empty_individual.Load.Area.Story.Leq = 115;
empty_individual.Load.Area.Story.LiveRR = 200;
empty_individual.Load.Area.Roof.Dead = 345;
empty_individual.Load.Area.Roof.Lr = 150;
empty_individual.Load.Area.Roof.Mass = 57.5;
empty_individual.Design.Ratio = [];
empty_individual.Design.RatioType = [];
empty_individual.Design.Location = [];
empty_individual.Design.ComboName = cell(1,10);
empty_individual.Design.ErrorSummary = cell(1,10);
empty_individual.Design.WarningSummary = cell(1,10);
DataBase = repmat(empty_individual, 3000, 1);

%% Disp Time

StartTime = clock;
disp ('*** START TIME ***')
disp (['Year    = ' num2str(StartTime(1))])
disp (['Month   = ' num2str(StartTime(2))])
disp (['Day     = ' num2str(StartTime(3))])
disp (['Hour    = ' num2str(StartTime(4))])
disp (['Minute  = ' num2str(StartTime(5))])
disp (['Seconds = ' num2str(StartTime(6))])

%%
%% GA
disp('Start GA alghoritm')

%% Problem Definition
disp('Problem Definition')

global NFE;
NFE = 0;
NFESAP = 0;

CostFunction = @(Gene) GeneWeight(Gene, Element, WSection);   % Cost Function

NumOfWSection = size(WSection,2);

NumOfAllElement = numel(Element);
nVar = NumOfAllElement;                 % Number of Decision Variables
VarSize=[1 nVar];                       % Decision Variables Matrix Size

%% GA Parameters
disp('GA Parameters')

MaxIt = 2000;      % Maximum Number of Iterations
nPop = 10;        % Population Size

pc = 0.6;                 % Crossover Percentage
nc = 2*round(pc*nPop/2);  % Number of Offsprings (Parnets)

pm = 0.2;                 % Mutation Percentage
nm = round(pm*nPop);      % Number of Mutants

mu = 0.1;       % Mutation Rate

ANSWER = questdlg('Choose selection method:','Genetic Algorith',...
    'Roulette Wheel','Tournament','Random','Roulette Wheel');

UseRouletteWheelSelection = strcmp(ANSWER,'Roulette Wheel');
UseTournamentSelection = strcmp(ANSWER,'Tournament');
UseRandomSelection = strcmp(ANSWER,'Random');

if UseRouletteWheelSelection
    beta = 8;         % Selection Pressure
end

if UseTournamentSelection
    TournamentSize = 3;   % Tournamnet Size
end

pause(0.1);
%% Initialization
disp('Initialization')

empty_individual.Position = [];
empty_individual.Cost = [];
empty_individual.CheckSF = [];

DataBase(2).ElementSectionNumber = 1;

pop = repmat(empty_individual,nPop,1);
TotalCAPVMin = WSection(1).Area/WSection(end).Area-1;

Z = NumberOfStory*NumOfElement;
BeamVarSize = [1 Z];
ColumnVarSize = [1 NumOfAllElement-Z];

CAPV(NumOfAllElement) = 0;

for i=1:nPop
    TotalCAPV = 1;
    CounterW = 0;
    while TotalCAPV ~= 0
        CounterW = CounterW + 1;
        pop(i).Position(1:Z) = randi([1 NumOfWSection-4],BeamVarSize);
        for m=Z+1:NumOfAllElement
            pop(i).Position(m) = randi([1 NumOfWSection]);
            Tabu = pop(i).Position(m);
            while (WSection(Tabu).Area < 0.03)
                pop(i).Position(m) = randi([1 NumOfWSection]);
                Tabu = pop(i).Position(m);
            end
        end
    
        disp(['Create Primary Population, Number Of Pop = ' num2str(i) ' , Number Of Try = ' num2str(CounterW)])
        Ret = SapModel.SetModelIsLocked(false());
        if Ret ~= 0
            disp(['! Error at SapModel.SetModelIsLocked Number = ' num2str(CounterW)]);
        end
        for k = 1:NumOfAllElement
            Ret = FrameObj.SetSection(...
                Element(k).Name, WSection(pop(1).Position(k)).Name);
            if Ret ~= 0
                disp(['! Error at FrameObj.SetSection NumOfAllElement = ' num2str(k)])
            end
        end
        Ret = Analyze.RunAnalysis();
        if Ret ~= 0
            disp('! Error at Analyze.RunAnalysis ')
        end
        Ret = DesignSteel.StartDesign();
        if Ret ~= 0
            disp('! Error at DesignSteel.StartDesign ')
        end
        Ret = SelectObj.All;
        if Ret ~= 0
            disp('! Error at SelectObj.All ')
        end
        [Ret ,DesignedNumberItems, DesignedFrameName, DesignedRatio, ...
            DesignedRatioType, DesignedLocation, DesignedComboName, ...
            DesignedErrorSummary, DesignedWarningSummary ] = ...
            DesignSteel.GetSummaryResults('', DesignedNumberItems, DesignedFrameName,...
            DesignedRatio, DesignedRatioType, DesignedLocation, DesignedComboName,...
            DesignedErrorSummary, DesignedWarningSummary, SAP2000v19.eItemType.SelectedObjects);
        if Ret ~= 0
            disp('! Error at DesignSteel.GetSummaryResults ')
        end
        NFESAP = NFESAP+1;

        DataBase(NFESAP).ElementSectionNumber = pop(1).Position;
        DataBase(NFESAP).BeamSectionNumber = pop(1).Position(1:Z);
        DataBase(NFESAP).ColumnSectionNumber = pop(1).Position(Z+1:end);
    
        for l = 1:DesignedNumberItems
        DataBase(NFESAP).Design.Ratio(l) = DesignedRatio(l);
        DataBase(NFESAP).Design.RatioType(l) = DesignedRatioType(l);
        DataBase(NFESAP).Design.Location(DesignedNumberItems) = DesignedLocation(l);
        DataBase(NFESAP).Design.ComboName{DesignedNumberItems} = DesignedLocation(l);
        DataBase(NFESAP).Design.ErrorSummary{DesignedNumberItems} = DesignedLocation(l);
        DataBase(NFESAP).Design.WarningSummary{DesignedNumberItems} = DesignedLocation(l);
        end
        
        Ret = SelectObj.ClearSelection;
        if Ret ~= 0
            disp('! Error at SelectObj.ClearSelection ')
        end
    
        for k = 1:NumOfAllElement
            CAPV(k) = max(DesignedRatio(k)-1, 0);
        end
        TotalCAPV = sum(CAPV(:));
        if TotalCAPV > 0
            TotalCAPV = max(TotalCAPV,TotalCAPVMin);
        end
        for k = 1:NumOfAllElement
            if (DesignedRatio(k)==0)
                TotalCAPV = max(TotalCAPV,TotalCAPVMin);
            end
        end
    end
end

for i=1:nPop   
    
    % Evaluation
    pop(i).Cost = CostFunction(pop(i).Position);
    
    % Safty Factor
    pop(i).CheckSF = 1;
    
end

% Sort Population
Costs = [pop.Cost];
[Costs, SortOrder] = sort(Costs);
pop = pop(SortOrder);

% Store Best Solution
BestSol = repmat(empty_individual,MaxIt,1);

% Array to Hold Best Cost Values
BestCost = zeros(MaxIt,1);

% Store Cost
WorstCost = pop(end).Cost;

% Array to Hold Number of Function Evaluations
nfe = zeros(MaxIt,1);
nfesap = zeros(MaxIt,1);

%% GA Main Loop
disp('GA Main Loop')

for it = 1:MaxIt
    
% Calculate Selection Probabilities
if UseRouletteWheelSelection
    P = exp(-beta*Costs/WorstCost);
    P = P/sum(P);
end
    
% Crossover
popc = repmat(empty_individual,nc/2,2);
for k = 1:nc/2
    % Select Parents Indices
    if UseRouletteWheelSelection
        i1 = RouletteWheelSelection(P);
        i2 = RouletteWheelSelection(P);
    end
    if UseTournamentSelection
        i1 = TournamentSelection(pop,TournamentSize);
        i2 = TournamentSelection(pop,TournamentSize);
    end
    if UseRandomSelection
        i1 = randi([1 nPop]);
        i2 = randi([1 nPop]);
    end

    % Select Parents
    p1 = pop(i1);
    p2 = pop(i2);
        
    % Apply Crossover
    [popc(k,1).Position, popc(k,2).Position] = ...
                            Crossover(p1.Position, p2.Position);
                        
    % Evaluate Offsprings
    popc(k,1).Cost = CostFunction(popc(k,1).Position);
    popc(k,2).Cost = CostFunction(popc(k,2).Position);
       
end
popc = popc(:);

% Mutation
if (rand < 0.1)
    popm = repmat(empty_individual,nm,1);

    for k = 1:nm
        % Select Parent
        i = randi([1 nPop]);
        p = pop(i);
        
        % Apply Mutation
        popm(k).Position = Mutate(p.Position, mu, NumOfWSection);
        
        % Evaluate Mutant
        popm(k).Cost = CostFunction(popm(k).Position);   
    end
    % Create Merged Population
    
    pop = [pop; popc; popm];
else
    pop = [pop; popc];
end   

     
% Sort Population
Costs = [pop.Cost];
[Costs, SortOrder] = sort(Costs);
pop = pop(SortOrder);

% Safty Factor
for k = 1:size(pop,1)
    if isempty(pop(k).CheckSF)
        pop(k).CheckSF = 0;
    end
end

% Update Worst Cost
WorstCost = max(WorstCost, pop(end).Cost);
CounterW = 0;
while (pop(1).CheckSF == 0)
    CounterW = CounterW + 1;
    disp(['Number Of Check Safty Factor = ' num2str(CounterW)])
    Ret = SapModel.SetModelIsLocked(false());
    if Ret ~= 0
        disp(['! Error at SapModel.SetModelIsLocked Number = ' num2str(CounterW)]);
    end
    for k = 1:NumOfAllElement
        Ret = FrameObj.SetSection(...
            Element(k).Name, WSection(pop(1).Position(k)).Name);
        if Ret ~= 0
            disp(['! Error at FrameObj.SetSection NumOfAllElement = ' num2str(k)])
        end
    end
    Ret = Analyze.RunAnalysis();
    if Ret ~= 0
        disp('! Error at Analyze.RunAnalysis ')
    end
    Ret = DesignSteel.StartDesign();
    if Ret ~= 0
        disp('! Error at DesignSteel.StartDesign ')
    end
    Ret = SelectObj.All;
    if Ret ~= 0
        disp('! Error at SelectObj.All ')
    end
    [Ret ,DesignedNumberItems, DesignedFrameName, DesignedRatio, ...
        DesignedRatioType, DesignedLocation, DesignedComboName, ...
        DesignedErrorSummary, DesignedWarningSummary ] = ...
        DesignSteel.GetSummaryResults('', DesignedNumberItems, DesignedFrameName,...
        DesignedRatio, DesignedRatioType, DesignedLocation, DesignedComboName,...
        DesignedErrorSummary, DesignedWarningSummary, SAP2000v19.eItemType.SelectedObjects);
    
    NFESAP = NFESAP+1;
    
    DataBase(NFESAP).ElementSectionNumber = pop(1).Position;
    DataBase(NFESAP).BeamSectionNumber = pop(1).Position(1:Z);
    DataBase(NFESAP).ColumnSectionNumber = pop(1).Position(Z+1:end);
    
    for i = 1:DesignedNumberItems
    DataBase(NFESAP).Design.Ratio(i) = DesignedRatio(i);
    DataBase(NFESAP).Design.RatioType(i) = DesignedRatioType(i);
    DataBase(NFESAP).Design.Location(DesignedNumberItems) = DesignedLocation(i);
    DataBase(NFESAP).Design.ComboName{DesignedNumberItems} = DesignedLocation(i);
    DataBase(NFESAP).Design.ErrorSummary{DesignedNumberItems} = DesignedLocation(i);
    DataBase(NFESAP).Design.WarningSummary{DesignedNumberItems} = DesignedLocation(i);
    end
    
    if Ret ~= 0
        disp('! Error at DesignSteel.GetSummaryResults ')
    end
    Ret = SelectObj.ClearSelection;
    if Ret ~= 0
        disp('! Error at SelectObj.ClearSelection ')
    end
    pop(1).CheckSF = 1;
    for k = 1:NumOfAllElement
        CAPV(k) = max(DesignedRatio(k)-1, 0);
    end
    TotalCAPV = sum(CAPV(:));
    if TotalCAPV > 0
        TotalCAPV = max(TotalCAPV,TotalCAPVMin);
    end
    for k = 1:NumOfAllElement
        if (DesignedRatio(k)==0)
            TotalCAPV = max(TotalCAPV,TotalCAPVMin);
        end
    end
    
    pop(1).Cost = pop(1).Cost*(1+TotalCAPV);
    
    Costs = [pop.Cost];
    [Costs, SortOrder] = sort(Costs);
    pop = pop(SortOrder);
end

% Truncation
pop = pop(1:nPop);
Costs = Costs(1:nPop);

% Store Best Solution Ever Found
BestSol = pop(1);
    
% Store Best Cost Ever Found
BestCost(it) = BestSol.Cost;
    
% Store NFE
nfe(it) = NFE;
nfesap(it) = NFESAP;
    
% Show Iteration Information
disp(['Iteration ' num2str(it) ': NFE = ' num2str(nfe(it)) ', NFESAP = ' num2str(nfesap(it)) ', Best Cost = ' num2str(BestCost(it))]);
%for i = 1:NumOfAllElement
%    if DesignedRatio(i) > 1
%    disp(num2str(i))
%    end
%end

end

%% Results

figure;
plot(nfe,BestCost,'LineWidth',2);
xlabel('NFE');
ylabel('Cost');

figure;
plot(nfesap,BestCost,'LineWidth',2);
xlabel('NEFSAP');
ylabel('Cost');

figure;
plot(BestCost,'LineWidth',2);
xlabel('IT');
ylabel('Cost');



%% close Sap2000

Ret = SapObject.ApplicationExit(false());

if Ret ~= 0
    disp('! Error at SapObject.ApplicationExit ')
end

File = 0;
PropMaterial = 0;
PropFrame = 0;
FrameObj = 0;
PointObj = 0;
View = 0;
LoadPatterns = 0;
Analyze = 0;
AnalysisResults = 0;
AnalysisResultsSetup = 0;
SapModel = 0;
SapObject = 0;

%% Disp Time

%% Disp Time

EndTime = clock;
disp ('*** END TIME ***')
disp (['Year    = ' num2str(EndTime(1))])
disp (['Month   = ' num2str(EndTime(2))])
disp (['Day     = ' num2str(EndTime(3))])
disp (['Hour    = ' num2str(EndTime(4))])
disp (['Minute  = ' num2str(EndTime(5))])
disp (['Seconds = ' num2str(EndTime(6))])

