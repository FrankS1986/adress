   M S B A M L   `   `   `  ����E  AAdressbuch, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:Adressbuch
Adressbuch  5      #    Adressbuch.Kontaktebearbeiten   5      W SPresentationCore, Version=6.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=6.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      *localclr-namespace:Adressbuch     5      
�������.��5
      �� -   5          ��Name$  Control��5
      .��5   
   �  � ColumnDefinitions 5   
   �� 5      �� 5       � RowDefinitions 5   
   �� 5      $	��170 �6   �� 5      .+�5   
   �� $��0��6   .+�5      �� -   5         ��Name$ bildkontakt��5      $	��200��6*   $	��200��66   $��
/hinzu.png��6C   6~   5      �� 5   
   $��1��6   $
��Left=�6%   .+�5      �� �� 5      .��6   	Vorname6%    �� 5      -   5         ��Name$ vorname��5      $
��11,2q�6'   $	��200��65   6C   5      �� 5      �� 5      .��6   
Nachname6&    �� 5      -   5        $
 name��5      $	��2,2q�6(   $	��200��65   6C   5       �� 5"      �� 5#      .��6   
Geburtstag6(    (  "System.Windows.Controls.DatePicker  5$      -   5          Name$ 
datePicker��5$      $	��0,2q�61   $	��200��6?   6M   5%      5(      �� 5*   
   $��1��6   $��2��6"   .+�5+      �� �� 5-      $��	Strasse: ��6   $��15(�6)   69    �� 5.      -   5        $ HinzufuegenStrasse��5.      $	��200��67   $��Center=�6C   $
��57,7q�6^   6n   50      �� 51      �� 52      $��Hausnummer: ��6   $��15(�60   6@    �� 53      -   5        $ HinzufuegenHausnummer��53      $	��200��6:   $��Center=�6F   $
��15,7q�6a   6�   54      �� 55      �� 56      $��Ort: ��6   $��15(�6)   69    �� 57      -   5        $ HinzufuegenOrt��57      $	��200��62   $��Center=�6>   $
��84,7q�6Y   6i   58      �� 59      �� 5:      $��Postleitzahl: ��6   $��15(�62   6B    �� 5;      -	   5        $ HinzufuegenPostleizahl��5;      $	��200��6:   $��Center=�6F   $
��30,7q�6a   6�   5<      �� 5=      �� 5>      $��	Telefon: ��6   $��15(�6-   6=    �� 5?      -
   5        $ HinzufuegenTelefon��5?      $	��200��66   $��Center=�6B   $
��58,7q�6]   6�   5@      �� 5A      �� 5B      $��	E-Mail:  ��6   $��15(�6,   6<    �� 5C      -   5        $ HinzufuegenEmail��5C      $	��200��64   $��Center=�6@   $
��59,7q�6[   6k    �� 5D      -   5         ��Name$ baestaetigen��5D      $��Bestätigen��62   $
��40,2q�6b   6r    �� 5E      -   5        $ 	abbrechen��5E      $��	Abbrechen��6(   $
��40,2q�6T   6d   5F      �� 5G      5I      5P      5R      5S      5T      