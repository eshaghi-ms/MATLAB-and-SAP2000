function SetMaterial()

%Material = xlsread(FileName);
%NumOfMaterial = size(Material,1);
%clear Material;
%[~,~,MPIsotropicData]   = xlsread(FileName,1,['A2:D' num2str(NumOfMaterial+1)]);
%[~,~,OSteelData]        = xlsread(FileName,3,['A2:D' num2str(NumOfMaterial+1)]);
%[~,~,OConcreteData]     = xlsread(FileName,4,['A2:D' num2str(NumOfMaterial+1)]);
%[~,~,WeightAndMassData] = xlsread(FileName,2,['A2:D' num2str(NumOfMaterial+1)]);

%% define material property

PropMaterial = ...
    NET.explicitCast(SapModel.PropMaterial,'SAP2000v19.cPropMaterial');

PropMaterial.SetMaterial('ST37Roller', SAP2000v19.eMatType.Steel);

PropMaterial.SetMaterial('C21', SAP2000v19.eMatType.Concrete);
PropMaterial.SetMaterial('C0' , SAP2000v19.eMatType.Concrete);

%% assign isotropic mechanical properties to material

PropMaterial.SetMPIsotropic('ST37Roller', 2E+10, 0.3, 1.17E-05);
PropMaterial.SetMPIsotropic('C21' , 2.495E+9, 0.15, 9.9E-06);
PropMaterial.SetMPIsotropic('C0'  , 2.495E+9, 0.15, 9.9E-06);

PropMaterial.SetOSteel_1('ST37Roller', 24E+6,...
                37E+6, 28.8E+6, 44.40E+6, 1, 1, 0.015, 0.11, 0.17, -0.1);
PropMaterial.SetOConcrete_1('C21', 21E+5, ...
                false, 0, 2, 2, 2.219E-03, 5.000E-03, -0.1, 0, 0);
PropMaterial.SetOConcrete_1('C0' , 21E+5, ...
                false, 0, 2, 2, 2.219E-03, 5.000E-03, -0.1, 0, 0);              

PropMaterial.SetWeightAndMass('ST37Roller', 1, 7850);
PropMaterial.SetWeightAndMass('C21', 1, 2500);
PropMaterial.SetWeightAndMass('C0' , 1, 0   );

PropMaterial.Delete('4000psi');
PropMaterial.Delete('A992Fy50');
end