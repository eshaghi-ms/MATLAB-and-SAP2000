function [W ,Sol] = CostFunction(Tour, Element, WSection)

global NFE;
if isempty(NFE)
    NFE=0;
end
NFE = NFE+1;

W = sum([Element(:).Length].*[WSection(Tour(:)).Area].*...
                                    [WSection(Tour(:)).WPerUnitV]);
                                
Sol.Check = 0;
end