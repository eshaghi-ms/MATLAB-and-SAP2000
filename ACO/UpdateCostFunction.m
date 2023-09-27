function [W, Sol]= UpdateCostFunction(Weight, DesignRatio, ViolationMin)
N = numel(DesignRatio);
Violation(N) = 0;
for i = 1:N
    Violation(i) = max(DesignRatio(i)-1,0);
end
TotalViolation = sum(Violation(i));
if TotalViolation ~= 0
    TotalViolation = max(TotalViolation, ViolationMin);
end

for k = 1:N
    if DesignRatio(k)==0
        TotalViolation = max(TotalViolation, ViolationMin);
	end        
end
                                
W = Weight*(1+TotalViolation);

Sol.Violation = Violation;
Sol.DesignRatio = DesignRatio;
Sol.TotalViolation = TotalViolation;
Sol.IsFeasible = (TotalViolation==0);
Sol.Weight = Weight;
Sol.Check = 1;

end

