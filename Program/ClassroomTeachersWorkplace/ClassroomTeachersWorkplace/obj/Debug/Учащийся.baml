   M S B A M L   `   `   `  ����U  QClassroomTeachersWorkplace, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullG(clr-namespace:ClassroomTeachersWorkplaceClassroomTeachersWorkplace  5   	   1    +ClassroomTeachersWorkplace.Учащийся   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   :local(clr-namespace:ClassroomTeachersWorkplace     5   	   _ [MaterialDesignThemes.Wpf, Version=4.9.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   -   5          � Title$-  'Добавление Учащегося��5	   	   $	��500��6&   $	��600��63    � Icon$ /MainIcon.png��6?    � WindowStartupLocation$ CenterScreen=�6T   $��Center=�6y   $��Center=�6�    � WindowStyle$ 
ToolWindow=�6�    � 
ResizeMode$ NoResize=�6�   .��5
      � .+�5   
   �� -   5         ��Name$ 	mainLabel��5      $��Добавление��6#   $��Center=�68   $��0,35,0,0q�6U   $	��Top=�6g   $��24(�6   $
��Bold"�6�   �� 5   
   -   5         ��Name$ txtName��5   !    �� 	MaxLength$ 50��6   $
��Left=�62   $��	35,80,0,0q�6M    �� TextWrapping$ NoWrap=�6`   )  #MaterialDesignThemes.Wpf.HintAssist	   Hint$.	 (Введите ФИО учащегося��6v   $	��Top=�6�   $	��250��6�    %  !MaterialDesignFloatingHintTextBox#��[  6�   6  �� 5   
   -   5        $ txtPhone��5   !   $ 13��6   $
��Left=�63   $��
35,122,0,0q�6N   $ NoWrap=�6b   $#	 Введите телефон��6x   $	��Top=�6�   $	��250��6�   #��[  6�   6  (  "System.Windows.Controls.DatePicker  5   
   -   5        
  Name$
 pickerDateOfBirtzd��5      $
��Left=�61   $��
35,175,0,0q�6L   $	��Top=�6`   $	��250��6x     MaterialDesignDatePicker#��[ 6�   $0	 *Выберите дату рождения��6�   �� 5   
   -   5         ��Name$ comboSocBit��5   .   $
��Left=�6   $	��250��6C   $G	 AВыберите социально бытовые условия��6O     MaterialDesignComboBox#��[ 6�   $��
35,218,0,0q�6�   $	��Top=�6�   �� 5   
   -   5        $ comboMatSot��5   .   $
��Left=�6   $	��250��6C   $B	 <Выберите материальное состояние��6O   #��[ 6�   $��
35,252,0,0q�6�   $	��Top=�6�   �� 5   
   -   5        $ comboTrudStady��5   .   $
��Left=�6   $	��250��6F   $7	 1Выберите трудности в учёбе��6R   #��[ 6�   $��
35,294,0,0q�6�   $	��Top=�6�   �� 5   
   -	   5        $ 
comboUchet��5   .   $
��Left=�6   $	��250��6B   $9	 3Выберите состояние на учёте��6N   #��[ 6�   $��
300,92,0,0q�6�   $	��Top=�6�   �� 5   
   -
   5        $ comboCommunication��5   .   $
��Left=�6   $	��250��6J   $8	 2Выберите нарушение общения��6V   #��[ 6�   $��300,134,0,0q�6�   $	��Top=�6�   �� 5   
   -   5         ��Name$ 
mainButton��5      $��Добавить��6%   $��Center=�68   $��	0,387,0,0q�6U   $	��Top=�6h   $	��250��6�    � Cursor$
 Hand��6�   �� 5   
   -   5        $ comboInvalid��5   .   $
��Left=�6   $	��250��6D   $/	 )Выберите инвалидность��6P   #��[ 6�   $��300,175,0,0q�6�   $	��Top=�6�   �� 5   
   -   5        $ comboGrupZdorov��5   .   $
��Left=�6   $	��250��6G   $4	 .Выберите группу здоровья��6S   #��[ 6�   $��300,218,0,0q�6�   $	��Top=�6�   �� 5   
   -   5        $ comboChernobol��5   .   $
��Left=�6   $	��250��6F   $	 ЧАЭС��6R   #��[ 6x   $��300,252,0,0q�6�   $	��Top=�6�   �� 5   
   -   5        $ txtHobbi��5   !   $ 50��6   $
��Left=�63   $��300,282,0,0q�6N   $ NoWrap=�6c   $'	 !Введите увлечения��6y   $	��Top=�6�   $	��250��6�   #��[  6�   �� 5   
   -   5        $ comboSirota��5   0   $��Center=�6   $	��250��6E   $	 Сирота��6Q   #��[ 6y   $��	0,334,0,0q�6�   $	��Top=�6�   5      5      