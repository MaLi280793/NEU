clear all
rng(2307) %Reproduzierbarkeit
Hauptpfad= ['C:\' ]; %Pfad muss angepasst werden

%Speicherort
eps=0;
tol=0;
Parameter=['RV_Parameter.xls'];
Rohdaten_BM=['Basis05_01.xlsx', 'Basis05_02.xlsx','Basis05_05.xlsx', 'Basis05_10.xlsx','Basis05_15.xlsx', 'Basis05_20.xlsx', 'Basis05_25.xlsx';...
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

  Zeilenvektor_Kennzahlen={'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' 'WC_ES/VaR_Komot' 'WC_ES/VaR_Komot' ''...
    '' 'Theo391_3' '' 'Theo3_1' 'WC_VaR' 'WC_ES' '' 'Bemerkung3_4' 'Bemerkung3_4' ''...
    'Koraollar3_2'  'Delta_WC' 'Delta_Quo_max' 'Delta_Quo_max' 'Delta_Quo_AV'};
 RohDaten_Frank='RohDaten_Frank.xlsx';
 %Haupttabellen


rowname1={'Dim05';'Dim10';'Dim15';'Dim20'};%muss bleiben
Bereich={'B2','B3','B4','B5'};%muss bleiben
%Lädt die Parameter der Randverteilungen
RV_Parameter = xlsread(Parameter);

alpha=[0.900 0.925 0.950 0.975 0.990 0.999];
Dim=[5 10 15 20];
Size=[1000 2000 5000 10000 15000 20000 25000];
N=1000;
 %preserveing of variables for speed in Abhängigkeit der Pfade
VaR_BC = zeros(N,1);
VaR_WC = zeros(N,1);
ES_BC = zeros(N,1);
ES_WC = zeros(N,1);
VaR_Como = zeros(N,1);
VaR_PF_com= zeros(N,1);
RM_Frank=zeros(length(alpha),4);
Mean=zeros(length(alpha),5);
VaR_Comontonic=zeros(length(alpha),2);

% Initialisiertung Haupttabellen
ErgebnisMatrix_Frank=zeros(28,48);


for u=1:length(Dim)

    for r=1:length(Size)
          
        D=Dim(1,u);
        %Lädt die Randverteilungen 
        RV_C = xlsread(Rohdaten_BM(u,r+(r-1)*15-(r-1):r*15)); 
        %Clayton Copula fit
        families1={'C'};
        [HACObjectC, fitLogC] = HACopulafit(RV_C, families1); 
        FrankR1=zeros(Size(1,r),D/5);
        FrankR2=zeros(Size(1,r),D/5);
        FrankR3=zeros(Size(1,r),D/5);
        FrankR4=zeros(Size(1,r),D/5);
        FrankR5=zeros(Size(1,r),D/5);
        


        RM_Frank = zeros(length(alpha),4);
        VaR_Comontonic= zeros(length(alpha),2);
        Mean= zeros(length(alpha),5);

        for l=1:length(alpha)

            for i=1:N

    
    

                Frank = HACopularnd(HACObjectC, Size(1,r));
                Ma_Frank=sort(Frank);


                for j=1:D/5
                    FrankR1(:,j)=icdf('gp',Ma_Frank(:,D-(D-j)),RV_Parameter(1,j),RV_Parameter(2,j),RV_Parameter(3,j));
                    FrankR2(:,j)=icdf('LogNormal',Ma_Frank(:,D-(D-(j+D/5))),RV_Parameter(4,j),RV_Parameter(5,j));
                    FrankR3(:,j)=icdf('Exponential',Ma_Frank(:,D-(D-(j+2*D/5))),RV_Parameter(6,j));
                    FrankR4(:,j)=icdf('wbl',Ma_Frank(:,D-(D/5*2-j)),RV_Parameter(7,j),RV_Parameter(8,j));
                    FrankR5(:,j)=icdf('Gamma',Ma_Frank(:,D-((D/5)-j)),RV_Parameter(9,j),RV_Parameter(10,j));
                end
    
                Frank_Matrix=[FrankR1 FrankR2 FrankR3 FrankR4 FrankR5];

                alpha_RM=alpha(1,l);
                % Calculate the index of the element corresponding to the alpha percentile
                idx = ceil(alpha_RM*Size(1,r));
                Frank_Matrix_BC = Frank_Matrix(1:idx,:);
                Frank_Matrix_WC = Frank_Matrix(idx+1:end,:);   
    
                %Komontonic VaR   
                VaR_Como(i,:)=sum(Frank_Matrix(idx+1,:));
                % summiert zuerst die Zeilen und sortiert anschließend VaR(X1+...Xn)
                Ma_sum=sort(sum(Frank_Matrix,2));
                
                


                %VaR ES/WC
                [VaR_BC(i,:)] = Rearrangement_Algorithmus_VaR_max(Frank_Matrix_BC, eps);
                [VaR_WC(i,:)]  = Rearrangment_Algorithmus_VaR_min(Frank_Matrix_WC, eps);

                %ES BC
                [ES_BC(i,:)] =  Rearrangement_Algorithmus_ES_max(Frank_Matrix, eps, Size(1,r),alpha_RM);    
                %ES WC
                Zeilensumme_Clayton=sum(Frank_Matrix,2);% ist schon sortiert (überprüfen)
               
                ES_WC(i,:)=sum(Zeilensumme_Clayton(floor((Size(1,r)*alpha_RM))+1:(Size(1,r))))/Size(1,r)/(1-alpha_RM);
             
            end

            RM_Frank(l,:)=[min(VaR_BC) max(VaR_WC) min(ES_BC) max(ES_WC)];
            Mean(l,:)=[mean(VaR_BC) mean(VaR_WC) mean(ES_BC) mean(ES_WC) mean(VaR_Como)];
            VaR_Comontonic(l,:)=[min(VaR_Como) max(VaR_Como)];
            
            
         
        end

        Delta_WC = RM_Frank(:,2)./Mean(:,5);
        
 
           

            
            
            
            % Berechnung der Spreads 
            Spread_VaR =RM_Frank(:,2)-RM_Frank(:,1);
            Spread_ES = RM_Frank(:,4)-RM_Frank(:,3);
            Spreads_Qut= Spread_VaR./Spread_ES; %passt
            
            
            % Berechnung der Differenz zwischen ES und VaR für Best und
            % Worst Case
            Diff_BC= RM_Frank(:,3)-RM_Frank(:,1);
            Diff_WC = RM_Frank(:,4)-RM_Frank(:,2);
           

            %KOnvergenz sachen
             Theo391_3=Delta_WC.*(VaR_Comontonic(:,2)./RM_Frank(:,4));% passt besser
             Theo391=RM_Frank(:,2)./RM_Frank(:,4);% passt
             

Kennzahlen=[Spreads_Qut NaN(6, 1) Theo391 NaN(6, 1)...
    ones(6,1) Delta_WC NaN(6, 1)...
    Theo391_3 NaN(6, 1); NaN(1, 9)];




    Konvergenz=[Diff_BC; NaN(1, 1); Diff_WC; NaN(1, 1)];
    

    % fast die Daten für die Haupttabellen für alle alpha zusammen
    Haupttabelle=[RM_Frank Spread_VaR Spread_ES  NaN(6, 1); NaN(1, 7)];


        
       % Speichert die Daten für alle Dimensionen
        %Haupttabelle
        ErgebnisMatrix_Frank(u*length(Haupttabelle(:,1))-(length(Haupttabelle(:,1))-1):u*length(Haupttabelle(:,1)),r*length(Haupttabelle(1,:))-(length(Haupttabelle(1,:))-1):r*length(Haupttabelle(1,:)))=[Haupttabelle];
        %Ergebnisse für Theoreme, Konvergenz (Dim) und andere Kennzahlen
        Kennzahlen_all(u*length(Kennzahlen(:,1))-(length(Kennzahlen(:,1))-1):u*length(Kennzahlen(:,1)),r*length(Kennzahlen(1,:))-(length(Kennzahlen(1,:))-1):r*length(Kennzahlen(1,:)))=[Kennzahlen];
        %Konvergenz durch Anzahl von Beobachtungen
        Konvergenz_all(u*length(Konvergenz(:,1))-(length(Konvergenz(:,1))-1):u*length(Konvergenz(:,1)),r*length(Konvergenz(1,:))-(length(Konvergenz(1,:))-1):r*length(Konvergenz(1,:)))=[Konvergenz];
       

    end
    %speichert die Parameter der Clayton Copula für 25000 Beobachtungen
CopulaParameter(u,1)=HACObjectC.Parameter;
    toc
end
%% Speichert Daten
    xlswrite([Hauptpfad RohDaten_Frank], CopulaParameter, 'Theta','B1');
    xlswrite([Hauptpfad RohDaten_Frank], rowname1, 'Theta','A1');
 xlswrite([Hauptpfad RohDaten_Frank], columnNames, 'Tabellen', 'A1');
 xlswrite([Hauptpfad RohDaten_Frank], rowname2, 'Tabellen', 'A2');
 xlswrite([Hauptpfad RohDaten_Frank], ErgebnisMatrix_Frank, 'Tabellen', 'B2');




 %Kennzahlen
xlswrite([Hauptpfad RohDaten_Frank], Kennzahlen_all, 'Kennzahlen', 'B2');
xlswrite([Hauptpfad RohDaten_Frank], Zeilenvektor_Kennzahlen, 'Kennzahlen', 'A1');
xlswrite([Hauptpfad RohDaten_Frank], rowname2, 'Kennzahlen', 'A2');



xlswrite([Hauptpfad RohDaten_Frank], Konvergenz_all, 'Konvergenz', 'B2');
xlswrite([Hauptpfad RohDaten_Frank], Row_konv, 'Konvergenz', 'A1');
xlswrite([Hauptpfad RohDaten_Frank], Colu_Konv, 'Konvergenz', 'A2');





