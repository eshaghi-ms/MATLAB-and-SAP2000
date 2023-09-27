function Weight = GeneWeight(Gene, Element, WSection)

global NFE;
if isempty(NFE)
    NFE=0;
end

NFE = NFE+1;

Weight = sum([Element(:).Length].*[WSection(Gene(:)).Area].*...
                                    [WSection(Gene(:)).WPerUnitV]);
end