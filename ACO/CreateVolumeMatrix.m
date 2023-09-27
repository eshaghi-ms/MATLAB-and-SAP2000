function V = CreateVolumeMatrix(Element, WSection)

NumOfElement  = numel(Element);
NumOfWSection = numel(WSection);
    
V = zeros(NumOfElement,NumOfWSection);
    
for i = 1:NumOfElement
    for j = 1:NumOfWSection
        V(i,j) = Element(i).Length*WSection(j).Area*WSection(j).WPerUnitV;
    end
end
end