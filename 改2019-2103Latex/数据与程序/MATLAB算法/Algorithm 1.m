close all; clear all; clc
getfilename=ls('20*.xl*'); 
%取目录下所有excel文件的文件名(.xls或.xlsx)，此时的excel文件数据均为无量纲化处理过后的标准化数据
filename = cellstr(getfilename);                                      %将字符型数组转换为cell型数组
num_of_files = length(filename);                                    %excel文件数目
for i=1:num_of_files                                                  %循环读入excel数据并存入结构体database中
    database(i) = struct('Name',filename(i),'Data',xlsread(filename{i}));
end
%注:该脚本将所有的excel数据读入到变量database中，database向量的每个元素为一结构，该结构由文件名Name和文件中的数据Data组成

Uijk=cat(3,database(1).Data,database(2).Data,database(3).Data,database(4).Data,database(5).Data,...
    database(6).Data,database(7).Data,database(8).Data,database(9).Data,...
    database(10).Data,database(11).Data,database(12).Data,database(13).Data) %构建三维数据

num_of_indicators=26 %指标数量
num_of_cities=35     %样本城市数量

cov=cat(3,cov(Uijk(:,:,1)),cov(Uijk(:,:,2)),cov(Uijk(:,:,3)),cov(Uijk(:,:,4)),cov(Uijk(:,:,5)),...
    cov(Uijk(:,:,6)),cov(Uijk(:,:,7)),cov(Uijk(:,:,8)),cov(Uijk(:,:,9)),...
    cov(Uijk(:,:,10)),cov(Uijk(:,:,11)),cov(Uijk(:,:,12)),cov(Uijk(:,:,13)))%构建三维协方差数组

R=cat(3,corrcoef(Uijk(:,:,1)),corrcoef(Uijk(:,:,2)),corrcoef(Uijk(:,:,3)),corrcoef(Uijk(:,:,4)),...
    corrcoef(Uijk(:,:,5)),corrcoef(Uijk(:,:,6)),corrcoef(Uijk(:,:,7)),corrcoef(Uijk(:,:,8)),corrcoef(Uijk(:,:,9)),...
    corrcoef(Uijk(:,:,10)),corrcoef(Uijk(:,:,11)),corrcoef(Uijk(:,:,12)),corrcoef(Uijk(:,:,13)))%构建三维相关系数数组

R(3,21,12)=0.9999
R(21,3,12)=0.9999                                              %为便于计算，将数据中非对角线为1的数据用0.9999代替
                                      
for k=1:num_of_files
    for j=1:num_of_indicators
        Wjk(1,j,k)=0
    end
end  %定义三维权重数组

for k=1:num_of_files
    for j=1:num_of_indicators
      rjt=R(j,:,k)  %提取第j行元素
      rj=R(:,j,k)  %提取第j列元素
      rj(find(rj==1))=[]
      rjt(find(rjt==1))=[]
      R1=R(:,:,k)
      R1(j,:)=[]%删除第j行元素
      R1(:,j)=[]%删除第j列元素
      R2=[R1 rj;rjt 1]
      INVR1=inv(R1)  %取逆
      Pj2(j)=rjt*INVR1*rj
    end
    Pj=sqrt(Pj2)  %求得复相关系数 
    pj=1./abs(Pj)
    for j=1:9
    Wjk(:,j,k)=[1./abs(Pj(j))]/sum(pj(1:9))   %归一化，求得在tk时刻的指标权重
    end
    for j=10:17
    Wjk(:,j,k)=[1./abs(Pj(j))]/sum(pj(10:17))   %归一化，求得在tk时刻的指标权重
    end
    for j=18:26
    Wjk(:,j,k)=[1./abs(Pj(j))]/sum(pj(18:26))   %归一化，求得在tk时刻的指标权重
    end
end
 

for i=1:num_of_cities
   for k=1:num_of_files
            Zih(i,k)=[sum(Uijk(i,1:9,k).* Wjk(1,1:9,k))]%住宅地产综合指数
    end
end
for i=1:num_of_cities
   for k=1:num_of_files 
            Zio(i,k)=[sum(Uijk(i,10:17,k).* Wjk(1,10:17,k))]%办公地产综合指数
    end
end
for i=1:num_of_cities
   for k=1:num_of_files
            Zic(i,k)=[sum(Uijk(i,18:26,k).* Wjk(1,18:26,k))]%商业地产综合指数
    end
end

for k=1:num_of_files
    for i=1:num_of_cities
        Wih(i,k)=[Zih(i,k)/sqrt(Zih(i,k)^2+Zio(i,k)^2+Zic(i,k)^2)] %住宅地产子系统的综合贡献Wih
    end
end

for k=1:num_of_files
    for i=1:num_of_cities
        Wio(i,k)=[Zio(i,k)/sqrt(Zih(i,k)^2+Zio(i,k)^2+Zic(i,k)^2)] %办公地产子系统的综合贡献Wio
    end
end

for k=1:num_of_files
    for i=1:num_of_cities  
        Wic(i,k)=[Zic(i,k)/sqrt(Zih(i,k)^2+Zio(i,k)^2+Zic(i,k)^2)] %商业地产子系统的综合贡献Wic
    end
end


for i=1:num_of_cities
    for k=1:num_of_files
        denominator=((Zih(i,k)+Zio(i,k)+Zic(i,k))/3)^3
        Ci(i,k)=(Zih(i,k)*Zio(i,k)*Zic(i,k)/denominator)^(1/3)%耦合度Ci
    end
end

    
for i=1:35
    for k=1:num_of_files
        Ti(i,k)=Wih(i,k)*Zih(i,k)+Wio(i,k)*Zio(i,k)+ Wic(i,k)*Zic(i,k)%协调程度（综合评价指数）Ti    
    end
end

for i=1:35
    for k=1:num_of_files
        Di(i,k)=(Ci(i,k)*Ti(i,k))^(1/2)%耦合协调度Di
    end
end



  %图1 35个大中城市住宅-商业-办公三维分布图
development1=Zih(:,13)
development2=Zio(:,13)
development3=Zic(:,13)
c=development3;%c表示对z轴进行着色
scatter3(development1,development2,development3,50,c)

L={'北京','天津','石家庄','太原','呼和浩特','沈阳','大连','长春','哈尔滨','上海',...
    '南京','杭州','宁波','合肥','福州','厦门','南昌','济南','青岛','郑州','武汉',...
    '长沙','广州','深圳','南宁','海口','重庆','成都','贵阳','昆明','西安',...
    '兰州','西宁','银川','乌鲁木齐',};
for i=1:35
text(development1(i),development2(i),development3(i),L(i))
end
       %%设置三维曲面x轴，y轴，z轴，标题对应内容及三个坐标轴的取值范围%%
       xlabel('X住宅地产');
       ylabel('Y办公地产');
       zlabel('Z商业地产');
       
figure_FontSize=12;
set(get(gca,'XLabel'),'FontSize',figure_FontSize,'Vertical','top');
set(get(gca,'YLabel'),'FontSize',figure_FontSize,'Vertical','middle');
set(findobj('FontSize',10),'FontSize',figure_FontSize);



figure;    %图2 2017年35个大中城市耦合度、协调度和耦合协调度折线图
 x=1:1:35;
 a=[0.894 ,0.985 ,0.958 ,0.994 ,0.964 ,0.965 ,0.958 ,0.986 ,0.971 ,0.974 ,0.960 ,0.987 ,0.987 ,0.974 ,0.996 ,0.961 ,0.982 ,0.992 ,0.860 ,0.998 ,0.974 ,0.877 ,0.981 ,...
     0.775 ,0.944 ,0.987 ,0.871 ,0.993 ,0.976 ,0.988 ,0.978 ,0.999 ,0.994 ,0.870 ,0.766 ]; %a数据y值
 b=[0.413,0.365,0.286 ,0.245 ,0.252 ,0.197 ,0.176 ,0.170 ,0.171 ,0.159 ,0.157 ,0.133 ,0.130 ,...
     0.119 ,0.105 ,0.103 ,0.100 ,0.095 ,0.102 ,0.082 ,0.081 ,0.089 ,0.073 ,0.087 ,0.065 ,0.056 ,0.063 ,0.054 ,0.054 ,0.045 ,0.043 ,0.039 ,0.033 ,0.027 ,0.025 ]; %b数据y值
 c=[0.608 ,0.600 ,0.524 ,0.493 ,0.493 ,0.436 ,0.410 ,0.409 ,0.407 ,0.394 ,0.388 ,0.362 ,0.358 ,0.341 ,0.323 ,0.315 ,0.313 ,0.307 ,0.296 ,0.286 ,0.280 ,0.279 ,0.267 ,...
     0.260 ,0.248 ,0.235 ,0.234 ,0.231 ,0.229 ,0.211 ,0.205 ,0.198 ,0.181 ,0.154 ,0.139]; %c数据y值

 plot(x,a,'-*b',x,b,'-or',x,c,'-xk'); %线性，颜色，标记
axis([0,35,0,1.2])  %确定x轴与y轴框图大小
set(gca,'XTick',1:35);
set(gca,'XTickLabel',{'重庆','上海','北京','成都','杭州','郑州','深圳','广州','武汉','西安','天津','长沙','青岛','南京','合肥','宁波','福州','昆明','厦门','南昌','长春',...
    '哈尔滨','济南','沈阳','南宁','石家庄','大连','贵阳','银川','海口','太原','乌鲁木齐','兰州','呼和浩特','西宁'})

set(gca,'YTick',[0:0.2:1]) %y轴范围0-1，间隔0.2
legend('耦合度','协调度','耦合协调度');   %右上角标注

ylabel('指标值') %y轴坐标描述

set(gca,'XTickLabel',[]); %将原坐标去掉
% xpoints和ypoints代表新坐标的位置
xpoints = get(gca,'XTick');
ypoints = 0*ones(1,35);
str = {'重庆','上海','北京','成都','杭州','郑州','深圳','广州','武汉','西安','天津','长沙','青岛','南京','合肥','宁波','福州','昆明','厦门','南昌','长春',...
    '哈尔滨','济南','沈阳','南宁','石家庄','大连','贵阳','银川','海口','太原','乌鲁木齐','兰州','呼和浩特','西宁'};
text(xpoints,ypoints,str,'HorizontalAlignment','right','rotation',90,'FontSize',13)



figure;    %图5 2017年主要城市群耦合协调度指数分布图
  x=1:1:23;
  a=[0.600,0.493,0.341,0.315,0.323,0,0.524,0.388,0.235,0,0.608,0.493,0,0.410,0.409,0,0.407,0.362,0.341,0,0.358,0.267,0]; %a数据y值
  n=0.4
 bar(a,n,'FaceColor',[0 0 1]) 
 
axis([0,23,0,0.8])  %确定x轴与y轴框图大小

set(gca,'XTick',1:23);
set(gca,'XTickLabel',{'上海','杭州','南京','宁波','合肥','  ','北京','天津','石家庄','  ','重庆','成都','  ','深圳' ,'广州',' ','武汉','长沙','南昌',' ','青岛','济南'})
set(gca,'YTick',[0:0.2:0.8]) %y轴范围0-1，间隔0.2
legend('耦合协调度','FontSize',14);   %右上角标注
ylabel('指标值','FontSize',14) %y轴坐标描述

set(gca,'XTickLabel',[]); %将原坐标去掉
% xpoints和ypoints代表新坐标的位置
xpoints = get(gca,'XTick');
ypoints = 0*ones(1,23);
str = {'上海','杭州','南京','宁波','合肥','  ','北京','天津','石家庄','  ','重庆','成都','  ','深圳' ,'广州',' ','武汉','长沙','南昌',' ','青岛','济南',' '};
text(xpoints,ypoints,str,'HorizontalAlignment','right','rotation',90,'FontSize',14)

for i = 1:length(a)
    text(a(i)+i-0.4,a(i),num2str(a(i)), 'HorizontalAlignment','center', 'VerticalAlignment','bottom')
end
