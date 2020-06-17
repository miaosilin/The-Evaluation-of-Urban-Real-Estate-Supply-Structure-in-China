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


