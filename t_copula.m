% Die Benchmark muss zuerst simuliert werden

clear all
rng(2307) %Reproduzierbarkeit
Hauptpfad= ['C: ...   \' ];% Pfad muss angepasst werden!!

%Speicherort

eps=0;
tol=0;
Parameter=['RV_Parameter.xls'];
%Namen der Excel-sheets für verschiede Dimensionen und Stichproben
Rohdaten=['Basis05_01.xlsx', 'Basis05_02.xlsx','Basis05_05.xlsx', 'Basis05_10.xlsx','Basis05_15.xlsx', 'Basis05_20.xlsx', 'Basis05_25.xlsx';...
    'Basis10_01.xlsx', 'Basis10_02.xlsx','Basis10_05.xlsx', 'Basis10_10.xlsx','Basis10_15.xlsx', 'Basis10_20.xlsx', 'Basis10_25.xlsx';...
    'Basis15_01.xlsx', 'Basis15_02.xlsx','Basis15_05.xlsx', 'Basis15_10.xlsx','Basis15_15.xlsx', 'Basis15_20.xlsx', 'Basis15_25.xlsx';...
    'Basis20_01.xlsx', 'Basis20_02.xlsx','Basis20_05.xlsx', 'Basis20_10.xlsx','Basis20_15.xlsx', 'Basis20_20.xlsx', 'Basis20_25.xlsx'];

columnNames = {'Dim05/1000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES',...
    '2000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES',...
    '5000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES',...
    '10000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES',...
    '15000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES',...
    '20000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES',...
    '25000','VaR_BC', 'VaR_WC', 'ES_BC',  'ES_WC','Spread_VaR','Spread_ES'};
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
    'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
    'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''};



rowname1={'Dim05';'Dim10';'Dim15';'Dim20'};%muss bleiben
Bereich={'B1','B2','B3','B4'};%muss bleiben
RohDaten_t={'Rohdaten_t'};
%Lädt die Parameter der Randverteilungen
RV_Parameter = xlsread(Parameter);

alpha=[0.900 0.925 0.950 0.975 0.990 0.999]; %Konfidenzniveaus
Dim=[5 10 15 20]; %Dimensionen
Size=[1000 2000 5000 10000 15000 20000 25000]; %Stichprobengröße
N=1000; %Anzahl der Pfade
%preserveing of variables for speed in Abhängigkeit der Pfade
VaR_BC = zeros(N,1);
VaR_WC = zeros(N,1);
ES_BC = zeros(N,1);
ES_WC = zeros(N,1);
VaR_Como = zeros(N,1);
VaR_PF_com= zeros(N,1);
RM_t=zeros(length(alpha),4);
Mean=zeros(length(alpha),5);
VaR_Comontonic=zeros(length(alpha),2);

% Initialisiertung Haupttabellen
ErgebnisMatrix_t=zeros(28,48);


for u=1:length(Dim) % Für jede Dimension eine neuer Durchlauf

    Rhohat_all=[]; % speichert alle Daten für jeweils eine Dimenison
    for r=1:length(Size) % Für jede Stichprobe ein neuer Durchlauf

        D=Dim(1,u);
        %Lädt die Randverteilungen
        RV_t = xlsread(Rohdaten(u,r+(r-1)*15-(r-1):r*15)); % lädt die entsprechende Benchmark Tabelle

        [Rhohat,Nuhat] = copulafit('t',RV_t); % Schätzt die Paremeter der t-Copula
        Nuhat_all(r,1)=Nuhat; % speichert die    Freiheitsgrade
        Rhohat_all=[Rhohat_all Rhohat NaN(length(Rhohat(1,:)), 1)];% Speichert Korrelationsmatrix

        % initialisiert Variablen abhängig von der Sampla Size und
        % Dimension
        tR1=zeros(Size(1,r),D/5);
        tR2=zeros(Size(1,r),D/5);
        tR3=zeros(Size(1,r),D/5);
        tR4=zeros(Size(1,r),D/5);
        tR5=zeros(Size(1,r),D/5);

        RM_t = zeros(length(alpha),4); % speicherplat für BC/WC ES und VaR
        VaR_Comontonic= zeros(length(alpha),2); %komontonische VaR
        Mean= zeros(length(alpha),5); % Speicherort für Erwartungswert aller RM

        for l=1:length(alpha) % Neuer Schleifen Durchlauf für jedes Konfidenzniveau

            for i=1:N
                tcopula = copularnd('t',Rhohat,Nuhat,Size(1,r)); %Simuliert t-Copula
                Ma_t=sort(tcopula); %sortiert die Matrix der Größe nach

                % Werte werden von Copula Skalar mit
                % Randverteilungsparameter zurück transformiert
                for j=1:D/5
                    tR1(:,j)=icdf('gp',Ma_t(:,D-(D-j)),RV_Parameter(1,j),RV_Parameter(2,j),RV_Parameter(3,j));
                    tR2(:,j)=icdf('LogNormal',Ma_t(:,D-(D-(j+D/5))),RV_Parameter(4,j),RV_Parameter(5,j));
                    tR3(:,j)=icdf('Exponential',Ma_t(:,D-(D-(j+2*D/5))),RV_Parameter(6,j));
                    tR4(:,j)=icdf('wbl',Ma_t(:,D-(D/5*2-j)),RV_Parameter(7,j),RV_Parameter(8,j));
                    tR5(:,j)=icdf('Gamma',Ma_t(:,D-((D/5)-j)),RV_Parameter(9,j),RV_Parameter(10,j));
                end

                t_Matrix=[tR1 tR2 tR3 tR4 tR5]; %speichert die transformierten Randverteilungen

                alpha_RM=alpha(1,l); % aktuelles Konfidenzniveau
                % Calculate the index of the element corresponding to the alpha percentile
                idx = ceil(alpha_RM*Size(1,r)); % Index abhängig von Konfidenzniveau und Stichprobe
                Gauss_Matrix_BC = t_Matrix(1:idx,:); % Definiert die Matrix für den BC-VaR
                Gauss_Matrix_WC = t_Matrix(idx+1:end,:);   % definiert die Matrix für WC-VaR

                %Komontonic VaR
                VaR_Como(i,:)=sum(t_Matrix(idx+1,:));% berechnet komontonische VaR
                % summiert zuerst die Zeilen und sortiert anschließend VaR(X1+...Xn)
                Ma_sum=sort(sum(t_Matrix,2));


                % speichert die RM für jeden Pfad (1000)
                %VaR ES/WC
                [VaR_BC(i,:)] = Rearrangement_Algorithmus_VaR_max(Gauss_Matrix_BC, eps); % berechnet mit RA den BC-VaR
                [VaR_WC(i,:)]  = Rearrangment_Algorithmus_VaR_min(Gauss_Matrix_WC, eps);% berechnet mit RA den WC-VaR

                %ES BC
                [ES_BC(i,:)] =  Rearrangement_Algorithmus_ES_max(t_Matrix, eps, Size(1,r),alpha_RM);   % berechnet mit RA den BC-VaR (nutzt gesamte Matrix)
                %ES WC
                Zeilensumme_Clayton=sum(t_Matrix,2);% summiertr die Zeilensumme der sortierten Matrix auf

                ES_WC(i,:)=sum(Zeilensumme_Clayton(floor((Size(1,r)*alpha_RM))+1:(Size(1,r))))/Size(1,r)/(1-alpha_RM); %Berechnet den WC-ES

            end

            RM_t(l,:)=[min(VaR_BC) max(VaR_WC) min(ES_BC) max(ES_WC)]; % speicher den Max bzw. Min wert von 1000 RM: BC-VaR, WC-VaR, BC-ES, WC-ES
            Mean(l,:)=[mean(VaR_BC) mean(VaR_WC) mean(ES_BC) mean(ES_WC) mean(VaR_Como)]; % berechnet den Erwarungswert der RM
            VaR_Comontonic(l,:)=[min(VaR_Como) max(VaR_Como)]; % wählt max und min von komo-VaR



        end

        Delta_WC = RM_t(:,2)./Mean(:,5); %berechnet WC-Delta

        %Verhältnis Worst Case ES/VaR
        Delta_Quo_max= RM_t(:,4)./VaR_Comontonic(:,2); % berechnet WC-ES/ komontonische-VaR

        % Berechnung der Spreads
        Spread_VaR =RM_t(:,2)-RM_t(:,1); % berechnet  VaR-Spread
        Spread_ES = RM_t(:,4)-RM_t(:,3); % berechnet ES-Spread
        Spreads_Qut= Spread_VaR./Spread_ES; %berechnet Quotient der Spreads


        % Berechnung der Differenz zwischen ES und VaR für Best und
        % Worst Case
        Diff_BC= RM_t(:,3)-RM_t(:,1); % berechnet Unterschied zwischen BC_ES-BC_VaR
        Diff_WC = RM_t(:,4)-RM_t(:,2); % berechnet Unterschied zwischen WC_ES-WC_VaR

        %KOnvergenz sachen
        Theo391_3=Delta_WC.*(VaR_Comontonic(:,2)./RM_t(:,4));% passt besser
        Theo391=RM_t(:,2)./RM_t(:,4);% passt


        % speichert die wichtigsten Kennzahlen aller Stichproben und
        % Konfidenzniveaus
        Kennzahlen=[Spreads_Qut NaN(6, 1) Theo391 NaN(6, 1)...
            ones(6,1) Delta_WC NaN(6, 1)...
            Theo391_3 Delta_Quo_max NaN(6, 1); NaN(1, 10)];

        Konvergenz=[Diff_BC; NaN(1, 1); Diff_WC; NaN(1, 1)];


        % fast die Daten für die Haupttabellen für alle alpha zusammen
        DATEN_C=[RM_t Spread_VaR Spread_ES  NaN(6, 1); NaN(1, 7)];

        % Speichert die Daten für alle Dimensionen
        %Haupttabelle
        ErgebnisMatrix_t(u*length(DATEN_C(:,1))-(length(DATEN_C(:,1))-1):u*length(DATEN_C(:,1)),r*length(DATEN_C(1,:))-(length(DATEN_C(1,:))-1):r*length(DATEN_C(1,:)))=[DATEN_C];
        %Ergebnisse für Theoreme, Konvergenz (Dim) und andere Kennzahlen
        Kennzahlen_all(u*length(Kennzahlen(:,1))-(length(Kennzahlen(:,1))-1):u*length(Kennzahlen(:,1)),r*length(Kennzahlen(1,:))-(length(Kennzahlen(1,:))-1):r*length(Kennzahlen(1,:)))=[Kennzahlen];
        %Konvergenz durch Anzahl von Beobachtungen
        Konvergenz_all(u*length(Konvergenz(:,1))-(length(Konvergenz(:,1))-1):u*length(Konvergenz(:,1)),r*length(Konvergenz(1,:))-(length(Konvergenz(1,:))-1):r*length(Konvergenz(1,:)))=[Konvergenz];
        Nuhat_all2(u*length(Nuhat_all(:,1))-(length(Nuhat_all(:,1))-1):u*length(Nuhat_all(:,1)),r*length(Nuhat_all(1,:))-(length(Nuhat_all(1,:))-1):r*length(Nuhat_all(1,:)))=[Nuhat_all];

    end
    %speichert die Parameter der Clayton Copula für 25000 Beobachtungen
    xlswrite([Hauptpfad RohDaten_t{1}], Rhohat_all, 'Theta',Bereich{u});

    toc
end
%% Speichert Daten alle in die gleiche Datei

xlswrite([Hauptpfad RohDaten_t{1}], columnNames, 'Tabellen', 'A1');
xlswrite([Hauptpfad RohDaten_t{1}], rowname2, 'Tabellen', 'A2');
xlswrite([Hauptpfad RohDaten_t{1}], ErgebnisMatrix_t, 'Tabellen', 'B2');
xlswrite([Hauptpfad RohDaten_t{1}], Nuhat_all2, 'Nu', 'A1');

%Kennzahlen
xlswrite([Hauptpfad RohDaten_t{1}], Kennzahlen_all, 'Kennzahlen', 'B2');
xlswrite([Hauptpfad RohDaten_t{1}], Zeilenvektor_Kennzahlen, 'Kennzahlen', 'A1');
xlswrite([Hauptpfad RohDaten_t{1}], rowname2, 'Kennzahlen', 'A2');


xlswrite([Hauptpfad RohDaten_t{1}], Konvergenz_all, 'Konvergenz', 'B2');
xlswrite([Hauptpfad RohDaten_t{1}], Row_konv, 'Konvergenz', 'A1');
xlswrite([Hauptpfad RohDaten_t{1}], Colu_Konv, 'Konvergenz', 'A2');





