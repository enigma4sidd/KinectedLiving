   M S B A M L   `   `   `  ���� T  PDiscreteGestureBasics-WPF, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullo<clr-namespace:Microsoft.Samples.Kinect.DiscreteGestureBasics.Microsoft.Samples.Kinect.DiscreteGestureBasics  ?    9Microsoft.Samples.Kinect.DiscreteGestureBasics.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Nlocal<clr-namespace:Microsoft.Samples.Kinect.DiscreteGestureBasics     -     � Title$  KINECTed LIVING��$	��650��$	��750�� � 	Resources    MediumGreyBrush �� Color  KinectPurpleBrush  KinectBlueBrush �� DataTypeF   @Microsoft.Samples.Kinect.DiscreteGestureBasics.GestureResultView � RowDefinitions �� StringFormat��%�  &	        &	      &	 *     (�� ?     *, +)�� $ 	#ff6e6e6e���� $ 	#ff52318f���� $ 	#ff00BCF2���� "  .R�� $
��Auto��$��5q����� *	BodyColor+ �� $
��Auto ��� �� $
��Auto ��� $
��Auto �.+��� $��Uniform=�$��1��$��5q����� *ImageSource+�� $��18(�$��Segoe UI)�$��SemiBold"�$��2��$��5 5 0 0q����� *
Detected+$ Gesture: {0}���� $��18(�$��Segoe UI)�$��SemiBold"�$��3��$��5 0 0 0q����� *
Confidence+$ Confidence: {0}��.��� $��	10 0 10 0q� �� $
��Auto ��� $��* ��� $
��Auto �.+��� $��0��$��Images\Logo.png��$
��Left=�$
��Fill=�$��32��$��81��$��0 10 0 5q��� $��0��$��	0,31,-1,0q�$��Right=�$	��Top=�#��[  $��Segoe UI)�$��25(�$
��Bold"�KINECTed LIVING�� $��0��$��Images\Status.png��$
��None=�$��Center=�$��0 0 0 5q�� -    �Name$ contentGrid��$��1�� � ColumnDefinitions �� $��1* ��� $��1* �.+���  �� Orientation$ Vertical=�$��1��$��3��$��
20 10 0 10q�$��Center=�$��Center=�.+�(� -   	 (�Name$	 kinectBodyViewbox��.���� $��UniformToFill=����� *ImageSource+�� $��Connect Microsoft Kinect��$
��Wrap=�$��5 10 5 5q�#��[  $��Segoe UI)�$��14(��� -   $��	statusBar��$��2��$��Stretch=�$��Bottom=�
�������#��[  .@��� ���� *
StatusText+