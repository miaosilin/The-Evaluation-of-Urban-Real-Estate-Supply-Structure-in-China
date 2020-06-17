close all; clear all; clc
getfilename=ls('20*.xl*'); 
%ȡĿ¼������excel�ļ����ļ���(.xls��.xlsx)����ʱ��excel�ļ����ݾ�Ϊ�����ٻ��������ı�׼������
filename = cellstr(getfilename);                                      %���ַ�������ת��Ϊcell������
num_of_files = length(filename);                                    %excel�ļ���Ŀ
for i=1:num_of_files                                                  %ѭ������excel���ݲ�����ṹ��database��
    database(i) = struct('Name',filename(i),'Data',xlsread(filename{i}));
end
%ע:�ýű������е�excel���ݶ��뵽����database�У�database������ÿ��Ԫ��Ϊһ�ṹ���ýṹ���ļ���Name���ļ��е�����Data���

Uijk=cat(3,database(1).Data,database(2).Data,database(3).Data,database(4).Data,database(5).Data,...
    database(6).Data,database(7).Data,database(8).Data,database(9).Data,...
    database(10).Data,database(11).Data,database(12).Data,database(13).Data) %������ά����

num_of_indicators=26 %ָ������
num_of_cities=35     %������������

cov=cat(3,cov(Uijk(:,:,1)),cov(Uijk(:,:,2)),cov(Uijk(:,:,3)),cov(Uijk(:,:,4)),cov(Uijk(:,:,5)),...
    cov(Uijk(:,:,6)),cov(Uijk(:,:,7)),cov(Uijk(:,:,8)),cov(Uijk(:,:,9)),...
    cov(Uijk(:,:,10)),cov(Uijk(:,:,11)),cov(Uijk(:,:,12)),cov(Uijk(:,:,13)))%������άЭ��������

R=cat(3,corrcoef(Uijk(:,:,1)),corrcoef(Uijk(:,:,2)),corrcoef(Uijk(:,:,3)),corrcoef(Uijk(:,:,4)),...
    corrcoef(Uijk(:,:,5)),corrcoef(Uijk(:,:,6)),corrcoef(Uijk(:,:,7)),corrcoef(Uijk(:,:,8)),corrcoef(Uijk(:,:,9)),...
    corrcoef(Uijk(:,:,10)),corrcoef(Uijk(:,:,11)),corrcoef(Uijk(:,:,12)),corrcoef(Uijk(:,:,13)))%������ά���ϵ������

R(3,21,12)=0.9999
R(21,3,12)=0.9999                                              %Ϊ���ڼ��㣬�������зǶԽ���Ϊ1��������0.9999����
                                      
for k=1:num_of_files
    for j=1:num_of_indicators
        Wjk(1,j,k)=0
    end
end  %������άȨ������

for k=1:num_of_files
    for j=1:num_of_indicators
      rjt=R(j,:,k)  %��ȡ��j��Ԫ��
      rj=R(:,j,k)  %��ȡ��j��Ԫ��
      rj(find(rj==1))=[]
      rjt(find(rjt==1))=[]
      R1=R(:,:,k)
      R1(j,:)=[]%ɾ����j��Ԫ��
      R1(:,j)=[]%ɾ����j��Ԫ��
      R2=[R1 rj;rjt 1]
      INVR1=inv(R1)  %ȡ��
      Pj2(j)=rjt*INVR1*rj
    end
    Pj=sqrt(Pj2)  %��ø����ϵ�� 
    pj=1./abs(Pj)
    for j=1:9
    Wjk(:,j,k)=[1./abs(Pj(j))]/sum(pj(1:9))   %��һ���������tkʱ�̵�ָ��Ȩ��
    end
    for j=10:17
    Wjk(:,j,k)=[1./abs(Pj(j))]/sum(pj(10:17))   %��һ���������tkʱ�̵�ָ��Ȩ��
    end
    for j=18:26
    Wjk(:,j,k)=[1./abs(Pj(j))]/sum(pj(18:26))   %��һ���������tkʱ�̵�ָ��Ȩ��
    end
end
 

for i=1:num_of_cities
   for k=1:num_of_files
            Zih(i,k)=[sum(Uijk(i,1:9,k).* Wjk(1,1:9,k))]%סլ�ز��ۺ�ָ��
    end
end
for i=1:num_of_cities
   for k=1:num_of_files 
            Zio(i,k)=[sum(Uijk(i,10:17,k).* Wjk(1,10:17,k))]%�칫�ز��ۺ�ָ��
    end
end
for i=1:num_of_cities
   for k=1:num_of_files
            Zic(i,k)=[sum(Uijk(i,18:26,k).* Wjk(1,18:26,k))]%��ҵ�ز��ۺ�ָ��
    end
end

for k=1:num_of_files
    for i=1:num_of_cities
        Wih(i,k)=[Zih(i,k)/sqrt(Zih(i,k)^2+Zio(i,k)^2+Zic(i,k)^2)] %סլ�ز���ϵͳ���ۺϹ���Wih
    end
end

for k=1:num_of_files
    for i=1:num_of_cities
        Wio(i,k)=[Zio(i,k)/sqrt(Zih(i,k)^2+Zio(i,k)^2+Zic(i,k)^2)] %�칫�ز���ϵͳ���ۺϹ���Wio
    end
end

for k=1:num_of_files
    for i=1:num_of_cities  
        Wic(i,k)=[Zic(i,k)/sqrt(Zih(i,k)^2+Zio(i,k)^2+Zic(i,k)^2)] %��ҵ�ز���ϵͳ���ۺϹ���Wic
    end
end


for i=1:num_of_cities
    for k=1:num_of_files
        denominator=((Zih(i,k)+Zio(i,k)+Zic(i,k))/3)^3
        Ci(i,k)=(Zih(i,k)*Zio(i,k)*Zic(i,k)/denominator)^(1/3)%��϶�Ci
    end
end

    
for i=1:35
    for k=1:num_of_files
        Ti(i,k)=Wih(i,k)*Zih(i,k)+Wio(i,k)*Zio(i,k)+ Wic(i,k)*Zic(i,k)%Э���̶ȣ��ۺ�����ָ����Ti    
    end
end

for i=1:35
    for k=1:num_of_files
        Di(i,k)=(Ci(i,k)*Ti(i,k))^(1/2)%���Э����Di
    end
end



  %ͼ1 35�����г���סլ-��ҵ-�칫��ά�ֲ�ͼ
development1=Zih(:,13)
development2=Zio(:,13)
development3=Zic(:,13)
c=development3;%c��ʾ��z�������ɫ
scatter3(development1,development2,development3,50,c)

L={'����','���','ʯ��ׯ','̫ԭ','���ͺ���','����','����','����','������','�Ϻ�',...
    '�Ͼ�','����','����','�Ϸ�','����','����','�ϲ�','����','�ൺ','֣��','�人',...
    '��ɳ','����','����','����','����','����','�ɶ�','����','����','����',...
    '����','����','����','��³ľ��',};
for i=1:35
text(development1(i),development2(i),development3(i),L(i))
end
       %%������ά����x�ᣬy�ᣬz�ᣬ�����Ӧ���ݼ������������ȡֵ��Χ%%
       xlabel('Xסլ�ز�');
       ylabel('Y�칫�ز�');
       zlabel('Z��ҵ�ز�');
       
figure_FontSize=12;
set(get(gca,'XLabel'),'FontSize',figure_FontSize,'Vertical','top');
set(get(gca,'YLabel'),'FontSize',figure_FontSize,'Vertical','middle');
set(findobj('FontSize',10),'FontSize',figure_FontSize);



figure;    %ͼ2 2017��35�����г�����϶ȡ�Э���Ⱥ����Э��������ͼ
 x=1:1:35;
 a=[0.894 ,0.985 ,0.958 ,0.994 ,0.964 ,0.965 ,0.958 ,0.986 ,0.971 ,0.974 ,0.960 ,0.987 ,0.987 ,0.974 ,0.996 ,0.961 ,0.982 ,0.992 ,0.860 ,0.998 ,0.974 ,0.877 ,0.981 ,...
     0.775 ,0.944 ,0.987 ,0.871 ,0.993 ,0.976 ,0.988 ,0.978 ,0.999 ,0.994 ,0.870 ,0.766 ]; %a����yֵ
 b=[0.413,0.365,0.286 ,0.245 ,0.252 ,0.197 ,0.176 ,0.170 ,0.171 ,0.159 ,0.157 ,0.133 ,0.130 ,...
     0.119 ,0.105 ,0.103 ,0.100 ,0.095 ,0.102 ,0.082 ,0.081 ,0.089 ,0.073 ,0.087 ,0.065 ,0.056 ,0.063 ,0.054 ,0.054 ,0.045 ,0.043 ,0.039 ,0.033 ,0.027 ,0.025 ]; %b����yֵ
 c=[0.608 ,0.600 ,0.524 ,0.493 ,0.493 ,0.436 ,0.410 ,0.409 ,0.407 ,0.394 ,0.388 ,0.362 ,0.358 ,0.341 ,0.323 ,0.315 ,0.313 ,0.307 ,0.296 ,0.286 ,0.280 ,0.279 ,0.267 ,...
     0.260 ,0.248 ,0.235 ,0.234 ,0.231 ,0.229 ,0.211 ,0.205 ,0.198 ,0.181 ,0.154 ,0.139]; %c����yֵ

 plot(x,a,'-*b',x,b,'-or',x,c,'-xk'); %���ԣ���ɫ�����
axis([0,35,0,1.2])  %ȷ��x����y���ͼ��С
set(gca,'XTick',1:35);
set(gca,'XTickLabel',{'����','�Ϻ�','����','�ɶ�','����','֣��','����','����','�人','����','���','��ɳ','�ൺ','�Ͼ�','�Ϸ�','����','����','����','����','�ϲ�','����',...
    '������','����','����','����','ʯ��ׯ','����','����','����','����','̫ԭ','��³ľ��','����','���ͺ���','����'})

set(gca,'YTick',[0:0.2:1]) %y�᷶Χ0-1�����0.2
legend('��϶�','Э����','���Э����');   %���ϽǱ�ע

ylabel('ָ��ֵ') %y����������

set(gca,'XTickLabel',[]); %��ԭ����ȥ��
% xpoints��ypoints�����������λ��
xpoints = get(gca,'XTick');
ypoints = 0*ones(1,35);
str = {'����','�Ϻ�','����','�ɶ�','����','֣��','����','����','�人','����','���','��ɳ','�ൺ','�Ͼ�','�Ϸ�','����','����','����','����','�ϲ�','����',...
    '������','����','����','����','ʯ��ׯ','����','����','����','����','̫ԭ','��³ľ��','����','���ͺ���','����'};
text(xpoints,ypoints,str,'HorizontalAlignment','right','rotation',90,'FontSize',13)



figure;    %ͼ5 2017����Ҫ����Ⱥ���Э����ָ���ֲ�ͼ
  x=1:1:23;
  a=[0.600,0.493,0.341,0.315,0.323,0,0.524,0.388,0.235,0,0.608,0.493,0,0.410,0.409,0,0.407,0.362,0.341,0,0.358,0.267,0]; %a����yֵ
  n=0.4
 bar(a,n,'FaceColor',[0 0 1]) 
 
axis([0,23,0,0.8])  %ȷ��x����y���ͼ��С

set(gca,'XTick',1:23);
set(gca,'XTickLabel',{'�Ϻ�','����','�Ͼ�','����','�Ϸ�','  ','����','���','ʯ��ׯ','  ','����','�ɶ�','  ','����' ,'����',' ','�人','��ɳ','�ϲ�',' ','�ൺ','����'})
set(gca,'YTick',[0:0.2:0.8]) %y�᷶Χ0-1�����0.2
legend('���Э����','FontSize',14);   %���ϽǱ�ע
ylabel('ָ��ֵ','FontSize',14) %y����������

set(gca,'XTickLabel',[]); %��ԭ����ȥ��
% xpoints��ypoints�����������λ��
xpoints = get(gca,'XTick');
ypoints = 0*ones(1,23);
str = {'�Ϻ�','����','�Ͼ�','����','�Ϸ�','  ','����','���','ʯ��ׯ','  ','����','�ɶ�','  ','����' ,'����',' ','�人','��ɳ','�ϲ�',' ','�ൺ','����',' '};
text(xpoints,ypoints,str,'HorizontalAlignment','right','rotation',90,'FontSize',14)

for i = 1:length(a)
    text(a(i)+i-0.4,a(i),num2str(a(i)), 'HorizontalAlignment','center', 'VerticalAlignment','bottom')
end
