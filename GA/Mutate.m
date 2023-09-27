function y = Mutate(x, mu, Num)

    nVar = numel(x);
    
    nmu=ceil(mu*nVar);
    
    j=randsample(nVar,nmu);
    
    y=x;
    for k = 1:nmu
        y(j(k)) = max(y(j(k))+1 ,randi([1 Num]));
    end
end

