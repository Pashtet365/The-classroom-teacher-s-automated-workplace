   M S B A M L   `   `   `  ����U  QClassroomTeachersWorkplace, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullG(clr-namespace:ClassroomTeachersWorkplaceClassroomTeachersWorkplace  5   	   /    )ClassroomTeachersWorkplace.Справки   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   :local(clr-namespace:ClassroomTeachersWorkplace     5   	   _ [MaterialDesignThemes.Wpf, Version=4.9.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   -   5          � Title$)  #Добавление справки��5	   	   $	��300��6$   $	��300��61    � Icon$ /MainIcon.png��6=    � WindowStartupLocation$ CenterScreen=�6R   $��Center=�6w   $��Center=�6�    � WindowStyle$ 
ToolWindow=�6�    � 
ResizeMode$ NoResize=�6�   .��5
      � .+�5   
   �� -   5         ��Name$ 	mainLabel��5      $��Добавление��6#   $��Center=�68   $��0,35,0,0q�6U   $	��Top=�6g   $��24(�6   $
��Bold"�6�   �� 5   
   -   5         ��Name$ comboStudent��5   0   $��Center=�6   $	��250��6F   )  #MaterialDesignThemes.Wpf.HintAssist   Hint$) #Выберите учащегося��6R      MaterialDesignComboBox#��[  6�   $��0,80,0,0q�6�   $	��Top=�6�   (  "System.Windows.Controls.DatePicker  5   
   -   5          Name$ pickerStartDate��5      $��Center=�6.   $��	0,110,0,0q�6K   $	��Top=�6^   $	��250��6v     MaterialDesignDatePicker#��[ 6�   $2 ,Выберите начальную дату��6�     5   
   -   5        $ pickerLastDate��5      $��Center=�6-   $��	0,142,0,0q�6J   $	��Top=�6]   $	��250��6u   #��[ 6�   $0 *Выберите конечную дату��6�   �� 5   
   -   5        	 ��Name$	 
mainButton��5      $��Добавить��6%   $��Center=�68   $��	0,202,0,0q�6U   $	��Top=�6h   $	��250��6�   
 � Cursor$

 Hand��6�   �� 5   
   -   5        $ comboVidSpravky��5   0   $��Center=�6   $	��250��6I   $, &Выберите вид справки��6U   #��[  6�   $��	0,172,0,0q�6�   $	��Top=�6�   5      5      