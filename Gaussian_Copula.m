clear all
% Die Benchmark muss zuerst simuliert werden
rng(2307) %Reproduzierbarkeit
Hauptpfad= ['C: ...   Master_Matlab\'];% Pfad muss angepasst werden!!

%Speicherort
% filename_RV=['RV_Dim05';'RV_Dim10';'RV_Dim15';'RV_Dim20'];
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
        'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
        'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
        'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
        'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
        'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''...
        'Dim05' 'SpreadQuot' '' 'Theo391(VaR/ES)' '' '1' 'WC_Delta' '' 'Theo391_3' 'Delta_Quo_max' ''};
RohDaten_Gauss='RohdatenGauss';

 %Haupttabellen


rowname1={'Dim05';'Dim10';'Dim15';'Dim20'};%muss bleiben
Bereich={'B1','B2','B3','B4'};%muss bleiben
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
RM_Gauss=zeros(length(alpha),4);
Mean=zeros(length(alpha),5);
VaR_Comontonic=zeros(length(alpha),2);

% Initialisiertung Haupttabellen
ErgebnisMatrix_Gauss=zeros(28,48);


for u=1:length(Dim)

    Rhohat_all=[];
    for r=1:length(Size)
          
        D=Dim(1,u);
        %Lädt die Randverteilungen 
        RV_Gauss = xlsread(Rohdaten(u,r+(r-1)*15-(r-1):r*15)); 
        %Clayton Copula fit
        RHOHATg = copulafit('Gaussian',RV_Gauss);
  
        Rhohat_all=[Rhohat_all RHOHATg NaN(length(RHOHATg(1,:)), 1)];
        GaussR1=zeros(Size(1,r),D/5);
        GaussR2=zeros(Size(1,r),D/5);
        GaussR3=zeros(Size(1,r),D/5);
        GaussR4=zeros(Size(1,r),D/5);
        GaussR5=zeros(Size(1,r),D/5);
        


        RM_Gauss = zeros(length(alpha),4);
        VaR_Comontonic= zeros(length(alpha),2);
        Mean= zeros(length(alpha),5);

        for l=1:length(alpha)

            for i=1:N

                Gauss_copula = copularnd('Gaussian',RHOHATg,Size(1,r));
                Ma_Gauss=sort(Gauss_copula);


                for j=1:D/5
                    GaussR1(:,j)=icdf('gp',Ma_Gauss(:,D-(D-j)),RV_Parameter(1,j),RV_Parameter(2,j),RV_Parameter(3,j));
                    GaussR2(:,j)=icdf('LogNormal',Ma_Gauss(:,D-(D-(j+D/5))),RV_Parameter(4,j),RV_Parameter(5,j));
                    GaussR3(:,j)=icdf('Exponential',Ma_Gauss(:,D-(D-(j+2*D/5))),RV_Parameter(6,j));
                    GaussR4(:,j)=icdf('wbl',Ma_Gauss(:,D-(D/5*2-j)),RV_Parameter(7,j),RV_Parameter(8,j));
                    GaussR5(:,j)=icdf('Gamma',Ma_Gauss(:,D-((D/5)-j)),RV_Parameter(9,j),RV_Parameter(10,j));
                end
    
                Gauss_Matrix=[GaussR1 GaussR2 GaussR3 GaussR4 GaussR5];

                alpha_RM=alpha(1,l);
                % Calculate the index of the element corresponding to the alpha percentile
                idx = ceil(alpha_RM*Size(1,r));
                Gauss_Matrix_BC = Gauss_Matrix(1:idx,:);
                Gauss_Matrix_WC = Gauss_Matrix(idx+1:end,:);   
    
                %Komontonic VaR   
                VaR_Como(i,:)=sum(Gauss_Matrix(idx+1,:));
                % summiert zuerst die Zeilen und sortiert anschließend VaR(X1+...Xn)
                Ma_sum=sort(sum(Gauss_Matrix,2));
                
                


                %VaR ES/WC
                [VaR_BC(i,:)] = Rearrangement_Algorithmus_VaR_max(Gauss_Matrix_BC, eps);
                [VaR_WC(i,:)]  = Rearrangment_Algorithmus_VaR_min(Gauss_Matrix_WC, eps);

                %ES BC
                [ES_BC(i,:)] =  Rearrangement_Algorithmus_ES_max(Gauss_Matrix, eps, Size(1,r),alpha_RM);    
                %ES WC
                Zeilensumme_Clayton=sum(Gauss_Matrix,2);% ist schon sortiert (überprüfen)
               
                ES_WC(i,:)=sum(Zeilensumme_Clayton(floor((Size(1,r)*alpha_RM))+1:(Size(1,r))))/Size(1,r)/(1-alpha_RM);
             
            end

            RM_Gauss(l,:)=[min(VaR_BC) max(VaR_WC) min(ES_BC) max(ES_WC)];
            Mean(l,:)=[mean(VaR_BC) mean(VaR_WC) mean(ES_BC) mean(ES_WC) mean(VaR_Como)];
            VaR_Comontonic(l,:)=[min(VaR_Como) max(VaR_Como)];
            
            
         
        end

        Delta_WC = RM_Gauss(:,2)./Mean(:,5);
        
            
        %Verhältnis Worst Case ES/VaR
            Delta_Quo_max= RM_Gauss(:,4)./VaR_Comontonic(:,2); %passt für alle alpha
           

            
            
            
            % Berechnung der Spreads 
            Spread_VaR =RM_Gauss(:,2)-RM_Gauss(:,1);
            Spread_ES = RM_Gauss(:,4)-RM_Gauss(:,3);
            Spreads_Qut= Spread_VaR./Spread_ES; %passt
            
            
            % Berechnung der Differenz zwischen ES und VaR für Best und
            % Worst Case
            Diff_BC= RM_Gauss(:,3)-RM_Gauss(:,1);
            Diff_WC = RM_Gauss(:,4)-RM_Gauss(:,2);
           

            %KOnvergenz sachen
             Theo391_3=Delta_WC.*(VaR_Comontonic(:,2)./RM_Gauss(:,4));% passt besser
             Theo391=RM_Gauss(:,2)./RM_Gauss(:,4);% passt
             


 Kennzahlen=[Spreads_Qut NaN(6, 1) Theo391 NaN(6, 1)...
    ones(6,1) Delta_WC NaN(6, 1)...
    Theo391_3 Delta_Quo_max NaN(6, 1); NaN(1, 10)];

    Konvergenz=[Diff_BC; NaN(1, 1); Diff_WC; NaN(1, 1)];
    

    % fast die Daten für die Haupttabellen für alle alpha zusammen
    DATEN_C=[RM_Gauss Spread_VaR Spread_ES  NaN(6, 1); NaN(1, 7)];


        
       % Speichert die Daten für alle Dimensionen
        %Haupttabelle
        ErgebnisMatrix_Gauss(u*length(DATEN_C(:,1))-(length(DATEN_C(:,1))-1):u*length(DATEN_C(:,1)),r*length(DATEN_C(1,:))-(length(DATEN_C(1,:))-1):r*length(DATEN_C(1,:)))=[DATEN_C];
        %Ergebnisse für Theoreme, Konvergenz (Dim) und andere Kennzahlen
        Kennzahlen_all(u*length(Kennzahlen(:,1))-(length(Kennzahlen(:,1))-1):u*length(Kennzahlen(:,1)),r*length(Kennzahlen(1,:))-(length(Kennzahlen(1,:))-1):r*length(Kennzahlen(1,:)))=[Kennzahlen];
        %Konvergenz durch Anzahl von Beobachtungen
        Konvergenz_all(u*length(Konvergenz(:,1))-(length(Konvergenz(:,1))-1):u*length(Konvergenz(:,1)),r*length(Konvergenz(1,:))-(length(Konvergenz(1,:))-1):r*length(Konvergenz(1,:)))=[Konvergenz];
       

    end
    %speichert die Parameter der Clayton Copula für 25000 Beobachtungen
xlswrite([Hauptpfad RohDaten_Gauss], Rhohat_all, 'Theta',Bereich{u});
    toc
end
%% Speichert Daten
    %xlswrite([Hauptpfad RohDaten_Gauss], CopulaParameter, 'Theta','B1');
    %xlswrite([Hauptpfad RohDaten_Gauss], rowname1, 'Theta','A1');
 xlswrite([Hauptpfad RohDaten_Gauss], columnNames, 'Tabellen', 'A1');
 xlswrite([Hauptpfad RohDaten_Gauss], rowname2, 'Tabellen', 'A2');
 xlswrite([Hauptpfad RohDaten_Gauss], ErgebnisMatrix_Gauss, 'Tabellen', 'B2');




 %Kennzahlen
xlswrite([Hauptpfad RohDaten_Gauss], Kennzahlen_all, 'Kennzahlen', 'B2');
xlswrite([Hauptpfad RohDaten_Gauss], Zeilenvektor_Kennzahlen, 'Kennzahlen', 'A1');
xlswrite([Hauptpfad RohDaten_Gauss], rowname2, 'Kennzahlen', 'A2');



xlswrite([Hauptpfad RohDaten_Gauss], Konvergenz_all, 'Konvergenz', 'B2');
xlswrite([Hauptpfad RohDaten_Gauss], Row_konv, 'Konvergenz', 'A1');
xlswrite([Hauptpfad RohDaten_Gauss], Colu_Konv, 'Konvergenz', 'A2');





