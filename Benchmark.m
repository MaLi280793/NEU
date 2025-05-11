clear all
%Benchmark muss zuerst simuliert werden
rng(2307) %Reproduzierbarkeit

Hauptpfad= ['C: ...   Master_Matlab\' ];% Muss ergänzt werden

% Für jede Sample Size und Dimension wird ein Exceldokument erstellt, dass
% als Grundlage für die Schätzung der anderen Copula Parameter dient
Basis=['Basis05_01.xlsx', 'Basis05_02.xlsx','Basis05_05.xlsx', 'Basis05_10.xlsx','Basis05_15.xlsx', 'Basis05_20.xlsx', 'Basis05_25.xlsx';...
    'Basis10_01.xlsx', 'Basis10_02.xlsx','Basis10_05.xlsx', 'Basis10_10.xlsx','Basis10_15.xlsx', 'Basis10_20.xlsx', 'Basis10_25.xlsx';...
    'Basis15_01.xlsx', 'Basis15_02.xlsx','Basis15_05.xlsx', 'Basis15_10.xlsx','Basis15_15.xlsx', 'Basis15_20.xlsx', 'Basis15_25.xlsx';...
    'Basis20_01.xlsx', 'Basis20_02.xlsx','Basis20_05.xlsx', 'Basis20_10.xlsx','Basis20_15.xlsx', 'Basis20_20.xlsx', 'Basis20_25.xlsx'];

columnNames = {'Dim05/1000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES',...
    '2000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES',...
    '5000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES',...
    '10000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES',...
    '15000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES',...
    '20000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES',...
    '25000','BM_VaR_BC', 'BM_VaR_WC', 'BM_ES_BC',  'BM_ES_WC','Spread_VaR','Spread_ES'};
rowname2={'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim10';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim15';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim20';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999'};

Row_konv={'Dim/Beob.' '1000' '2000' '5000' '10000' '15000' '20000' '25000'};

Colu_Konv={'Dim05 BC_Konv'; '0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim05 WC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim10 BC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim10 WC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim15 BC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim15 WC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim20 BC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999';...
    'Dim20 WC_Konv';'0.90';'0.925';'0.950';'0.975';'0.990';'0.999'};
colu_moments={'Dim05' 'VaR_BC' 'VaR_WC' 'ES_BC2' 'ES_WC' 'VaR_Como_BM' 'VaR_PF_com'};

Zeilenvektor_Kennzahlen={'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''};

%Haupttabellen


rowname1={'Dim05';'Dim10';'Dim15';'Dim20'};%muss bleiben
Bereich={'B1','B2','B3','B4'};%muss bleiben
%Konfidenzniveau
alpha=[0.900 0.925 0.950 0.975 0.990 0.999];
%Dimensionen
Dim=[5 10 15 20];
%sample size
Size=[1000 2000 5000 10000 15000 20000 25000];
% Anzahl der Pfade
N=1000;
tol=0;
%Initialisierung der Gemeinsamenverteilung (alle Randverteilunge haben eine Abhängigkeit von Theta=5)
C = cell(1,4);
C{1} =  {{'C', 5},1, 2, 3, 4, 5};
C{2} =  {{'C', 5},1, 2, 3, 4, 5, 6, 7, 8 ,9, 10};
C{3} = {{'C', 5},1, 2, 3, 4, 5, 6, 7, 8 ,9, 10, 11 , 12, 13, 14, 15};
C{4} =  {{'C', 5},1, 2, 3, 4, 5, 6, 7, 8 ,9, 10, 11 , 12, 13, 14, 15, 16, 17, 18, 19, 20};


%preserveing of variables for speed in Abhängigkeit der Pfade
VaR_BC=zeros(N,1);
VaR_WC=zeros(N,1);
ES_BC=zeros(N,1);
ES_WC=zeros(N,1);
VaR_Como_BM = zeros(N,1);

RM_BM=zeros(length(alpha),4);
Var_BM=zeros(length(alpha),6);
MEAN_BM=zeros(length(alpha),5);
VaR_Comontonic_BM=zeros(length(alpha),2);

% Initialisiertung Haupttabellen
ErgebnisMatrix=zeros(28,48);

% Left die Parameter für die Randverteilungen fest
%1.Randverteilung
% Pareto-Verteilung
% Zufällige aber fixe Wahl der Parameter der modifizierten Pareto-Verteilung
Pareto_k =  0.1 + (0.5-0.1).*rand(1,4); % shape parameter Formparameter im Intervall [1,3]
Pareto_sigma = 2 + (4-2).*rand(1,4); %scale parameter Formparameter im Intervall [0.5,2]
Pareto_mu =  0.00001+ (0.8-0.00001).*rand(1,4); %μ location parameter Skalenparameter im Intervall [0.1,1]

%2.Randverteilung
% Lognormal-Verteilung
LogNor_mu=(2.2).*rand(1,4);%Erwartungswert zufällig aus dem Intervall [0,2]
LogNor_sigma=1.5.*rand(1,4);%Sigam im Intervall [0.5,1.2] zufällig wählen

%3.Randverteilung
% Exponential-Verteilung
Expo_mu=15 + (20-15).*rand(1,4); %Lambda wird im Intervall [0.1,2] zufällig gewählt

%4.Randverteilung
% Weibull-Verteilung
Weibull_alpha =20+(20-15).*rand(1,4); %Lambda wird im Intervall [0.1,10] zufällig gewählt
Weibull_beta = 0.7 + (2-0.7).*rand(1,4); % Formparameter im Intervall [0.5,10]

%5.Randverteilung
% Gamma-Verteilung
Gamma_alpha = 0.8+ (2-0.8).*rand(1,4); %alpha wird im Intervall [0.1,5] zufällig gewählt
Gamma_beta = 15 + (20-15).*rand(1,4); %beta wird im Intervall [0.1,5] zufällig gewählt


% Speichert die Parameter der Randverteilungen als Exceltabelle
column_Names={'Pareto_k', 'Pareto_sigma', 'Pareto_mu', 'LogNor_mu', 'LogNor_sigma', 'Expo_mu', 'Weibull_alpha', 'Weibull_beta','Gamma_alpha', 'Gamma_beta'};
xlswrite([Hauptpfad 'RV_Parameter'], rowname1'', 'B1');
xlswrite([Hauptpfad 'RV_Parameter'], column_Names', 'A2:A11');
xlswrite([Hauptpfad 'RV_Parameter'], [Pareto_k; Pareto_sigma; Pareto_mu; LogNor_mu; LogNor_sigma; Expo_mu; Weibull_alpha; Weibull_beta; Gamma_alpha; Gamma_beta],'B2:E11');



% Schleife für die Dimensionen
for u=1:length(Dim)
    %Schleife für Sample Size
    for r=1:length(Size)

        % wählt aktuelle Sample Size
        sz = [1 Size(1,r)];
        % wählt die aktuelle Dimension
        D=Dim(1,u);
        % Generiert die Benchmark
        myHAC = HACopula();
        cell2tree(myHAC, C{u});
        BenchMark=HACopularnd(myHAC ,Size(1,r));

        % Speichert die Benchmark für verschiedenen Sample Size und Dimension in
        % Exceltabellen
        xlswrite([Hauptpfad Basis(u,r+(r-1)*15-(r-1):r*15)], BenchMark);

        % speichert die Variablen für die Randverteilungen, in Abhängigkeit
        % der Sample Size und Dimension, vor
        BM_copula=zeros(Size(1,r),D);
        BM1=zeros(Size(1,r),D/5);
        BM2=zeros(Size(1,r),D/5);
        BM3=zeros(Size(1,r),D/5);
        BM4=zeros(Size(1,r),D/5);
        BM5=zeros(Size(1,r),D/5);


        % Transformiert die Daten mithilfe der Randverteilungsparameter von
        % der Copula auf die "normale" Skala
        for j=1:D/5
            BM1(:,j)=icdf('gp',BenchMark(:,D-(D-j)),Pareto_k(1,j),Pareto_sigma(1,j),Pareto_mu(1,j));
            BM2(:,j)=icdf('LogNormal',BenchMark(:,D-(D-(j+D/5))),LogNor_mu(1,j),LogNor_sigma(1,j));
            BM3(:,j)=icdf('Exponential',BenchMark(:,D-(D-(j+2*D/5))),Expo_mu(1,j));
            BM4(:,j)=icdf('wbl',BenchMark(:,D-(D/5*2-j)),Weibull_alpha(1,j),Weibull_beta(1,j));
            BM5(:,j)=icdf('Gamma',BenchMark(:,D-((D/5)-j)),Gamma_alpha(1,j),Gamma_beta(1,j));
        end
        BM=[BM1 BM2 BM3 BM4 BM5];
        % sortiert die Randverteilungen aufsteigend
        BM_Sort=sort(BM);








        % Schleife für verschiede Alpha Niveaus
        for l=1:length(alpha)

            % Schleife Anzhal der Pfade (generiert N Pfade)
            for i=1:N

                % simulating Variables
                BM_copula=HACopularnd(myHAC ,Size(1,r));
                % Transformiert die Verteilungen von der Copula Skala
                % mit Hilfe der Randverteilungsparameter
                for j=1:D/5
                    BM1(:,j)=icdf('gp',BM_copula(:,D-(D-j)),Pareto_k(1,j),Pareto_sigma(1,j),Pareto_mu(1,j));
                    BM2(:,j)=icdf('LogNormal',BM_copula(:,D-(D-(j+D/5))),LogNor_mu(1,j),LogNor_sigma(1,j));
                    BM3(:,j)=icdf('Exponential',BM_copula(:,D-(D-(j+2*D/5))),Expo_mu(1,j));
                    BM4(:,j)=icdf('wbl',BM_copula(:,D-(D/5*2-j)),Weibull_alpha(1,j),Weibull_beta(1,j));
                    BM5(:,j)=icdf('Gamma',BM_copula(:,D-((D/5)-j)),Gamma_alpha(1,j),Gamma_beta(1,j));
                end


                Ma=[BM1 BM2 BM3 BM4 BM5];
                BM=sort(Ma);
                % Wählt den aktuellen alpha Wert
                alpha_RM=alpha(1,l);

                % berechnet den Index für die Qunatile in Abhängigkeit
                % von Alpha
                idx = ceil(alpha_RM*Size(1,r));

                % Speichert die Risikomaße für jeden Pfad
                %Komontonic VaR
                VaR_Como_BM(i,:)=sum(BM(idx+1,:));
                % summiert zuerst die Zeilen und sortiert anschließend VaR(X1+...Xn)
                Ma_sum=sort(sum(Ma,2));

                % Matrix für den Worst Case VaR
                BM_Matrix_WC=BM(idx+1:end,:);
                % Matrix für den Best Case VaR und ES
                BM_Matrix_BC=BM(1:idx,:);



                % Nutzt den improved Rearrangealgorithm für die Berechnung des BC und WC VaR und des BC ES
                % Best Case Value at Risk
                [VaR_BC(i,:)] = Rearrangement_Algorithmus_VaR_max(BM_Matrix_BC, tol);

                %WC_VaR
                [VaR_WC(i,:)]  = Rearrangment_Algorithmus_VaR_min(BM_Matrix_WC, tol);

                %BC ES %Hier muss N angepasst werden im original code
                [ES_BC(i,:),ZZZ2] = Rearrangement_Algorithmus_ES_max(BM, tol, Size(1,r), alpha_RM);

                % Berechnung des WC ES
                %WC ES
                % Berechnung der Zeilensumme
                Zeilensumme_BM=sum(BM,2);

                % Berechnung des WC ES
                ES_WC(i,:)=sum(Zeilensumme_BM(floor((Size(1,r)*alpha_RM))+1:(Size(1,r))))/Size(1,r)/(1-alpha_RM);



            end
            % speichert die RM für jedes Konfidenzniveau
            RM_BM(l,:)=[min(VaR_BC) max(VaR_WC) min(ES_BC) max(ES_WC)];
            MEAN_BM(l,:)=[mean(VaR_BC) mean(VaR_WC) mean(ES_BC) mean(ES_WC) mean(VaR_Como_BM)];
            VaR_Comontonic_BM(l,:)=[min(VaR_Como_BM) max(VaR_Como_BM)];

        end
        Delta_WC = RM_BM(:,2)./MEAN_BM(:,5); %WC_VaR/mean_Como_VaR
        %Verhältnis Worst Case ES/VaR
        Delta_Quo_max= RM_BM(:,4)./VaR_Comontonic_BM(:,2); %WC_ES/Com_VaR

        % Berechnung der Spreads
        Spread_VaR =RM_BM(:,2)-RM_BM(:,1); % WC_VaR-BC_VaR
        Spread_ES = RM_BM(:,4)-RM_BM(:,3); % WC_ES-BC_ES
        Spreads_Qut= Spread_VaR./Spread_ES; % (WC_VaR-BC_VaR)/(WC_ES-BC_ES)


        % Berechnung der Differenz zwischen ES und VaR für Best und
        % Worst Case
        Diff_BC= RM_BM(:,3)-RM_BM(:,1); %BC_ES-BC_VaR
        Diff_WC = RM_BM(:,4)-RM_BM(:,2); %WC_ES-WC_VaR


        %KOnvergenz sachen
        Theo391_3=Delta_WC.*(VaR_Comontonic_BM(:,2)./RM_BM(:,4)); % WC_Delta*max_Com_VaR/WC_ES


        Theo391=RM_BM(:,2)./RM_BM(:,4); % WC_ES/WC_VaR


        Kennzahlen=[Spreads_Qut NaN(6, 1) Theo391 NaN(6, 1)...
            ones(6,1) Delta_WC NaN(6, 1)...
            Theo391_3 Delta_Quo_max NaN(6, 1); NaN(1, 10)];

        Konvergenz=[Diff_BC; NaN(1, 1); Diff_WC; NaN(1, 1)];


        % fast die Daten für die Haupttabellen für alle alpha zusammen
        DATEN=[RM_BM Spread_VaR Spread_ES  NaN(6, 1); NaN(1, 7)];



        % Speichert die Daten für alle Dimensionen
        %Haupttabelle
        ErgebnisMatrix(u*length(DATEN(:,1))-(length(DATEN(:,1))-1):u*length(DATEN(:,1)),r*length(DATEN(1,:))-(length(DATEN(1,:))-1):r*length(DATEN(1,:)))=[DATEN];
        %Ergebnisse für Theoreme, Konvergenz (Dim) und andere Kennzahlen
        KENN_ZAHLEN(u*length(Kennzahlen(:,1))-(length(Kennzahlen(:,1))-1):u*length(Kennzahlen(:,1)),r*length(Kennzahlen(1,:))-(length(Kennzahlen(1,:))-1):r*length(Kennzahlen(1,:)))=[Kennzahlen];
        %Konvergenz durch Anzahl von Beobachtungen
        Konvergenz_all(u*length(Konvergenz(:,1))-(length(Konvergenz(:,1))-1):u*length(Konvergenz(:,1)),r*length(Konvergenz(1,:))-(length(Konvergenz(1,:))-1):r*length(Konvergenz(1,:)))=[Konvergenz];

    end
    %speichert die Parameter der Copula für 25000 Beobachtungen
    xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], myHAC.Parameter, 'Theta',Bereich{u});
    xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], rowname1, 'Theta','A1');


end
%% Speichert Daten

xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], columnNames, 'Tabellen', 'A1');
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], rowname2, 'Tabellen', 'A2');
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], ErgebnisMatrix, 'Tabellen', 'B2');


%Kennzahlen
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], KENN_ZAHLEN, 'Kennzahlen', 'B2');
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], Zeilenvektor_Kennzahlen, 'Kennzahlen', 'A1');
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], rowname2, 'Kennzahlen', 'A2');


xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], Konvergenz_all, 'Konvergenz', 'B2');
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], Row_konv, 'Konvergenz', 'A1');
xlswrite([Hauptpfad 'Rohdaten_BM.xlsx'], Colu_Konv, 'Konvergenz', 'A2');



