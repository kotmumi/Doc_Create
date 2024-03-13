using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using System.Windows;
using Doc_Create;
using System.Windows.Forms;
using InvAddIn.Properties;

namespace InvAddIn
{
    class MyButton
    {
        private Inventor.Application _inventor;
        private ButtonDefinition _settingsButton;
        private ButtonDefinition _settingsButtonDxfPart;

        public MyButton(Inventor.Application inventor)
        {
            _inventor = inventor;
            SetupButtonDefinition();
            SetupButtonDefinitionDxfPart();
            AddButtonDefinitionToRibbon();
        }

        private void SetupButtonDefinition()
        {
            ControlDefinitions conDefs = _inventor.CommandManager.ControlDefinitions;
            _settingsButton = conDefs.AddButtonDefinition(
                "Создать DXF файлы",
                "MyButton DisplayName",
                CommandTypesEnum.kEditMaskCmdType,
                Guid.NewGuid().ToString(),
                "Функция создает DXF файлы ",//деталей для лазерной резки всех листовых деталей из активной сборки",
                " Функция создает DXF файлы деталей для лазерной резки всех листовых деталей из активной сборки");
            _settingsButton.OnExecute += MyButton_OnExecute;
            _settingsButton.StandardIcon = PictureDispConverter.ToIPictureDisp(Resources.dxf_file16x16);
            _settingsButton.LargeIcon = PictureDispConverter.ToIPictureDisp(Resources.dxf_file32x32);
        }
    private void SetupButtonDefinitionDxfPart()
    {
            ControlDefinitions conDefs3 = _inventor.CommandManager.ControlDefinitions;
            //создание кнопки экспорта PDF
            _settingsButtonDxfPart = conDefs3.AddButtonDefinition(
                "Создать DXF файл",
                "MyButton DisplayName3",
                CommandTypesEnum.kEditMaskCmdType,
                Guid.NewGuid().ToString(),
                "Функция создает DXF файл.",//деталей для лазерной резки всех листовых деталей из активной сборки",
                " Функция создает DXF файл активной детали для лазерной резки.");
            _settingsButtonDxfPart.OnExecute += MyButton_OnExecuteDxfPart;
            _settingsButtonDxfPart.StandardIcon = PictureDispConverter.ToIPictureDisp(Resources.dxf_file16x16);
           _settingsButtonDxfPart.LargeIcon = PictureDispConverter.ToIPictureDisp(Resources.dxf_file32x32);
    }
        private void AddButtonDefinitionToRibbon()
        {
            //добавление кнопки экспорта DXF
            Ribbon ribbon = _inventor.UserInterfaceManager.Ribbons["Assembly"];
            RibbonTab ribbonTab = ribbon.RibbonTabs.Add("Савушкин продукт", "Савушкин продукт", Guid.NewGuid().ToString());
            RibbonPanel ribbonPanel = ribbonTab.RibbonPanels.Add("Документация", "Документация", Guid.NewGuid().ToString());
            ribbonPanel.CommandControls.AddButton(_settingsButton,true);

            Ribbon ribbonPart = _inventor.UserInterfaceManager.Ribbons["Part"];
            RibbonTab ribbonTabPart = ribbonPart.RibbonTabs.Add("Савушкин продукт", "Савушкин продуктPart", Guid.NewGuid().ToString());
            RibbonPanel ribbonPanelPart = ribbonTabPart.RibbonPanels.Add("Документация", "ДокументацияPart", Guid.NewGuid().ToString());
            ribbonPanelPart.CommandControls.AddButton(_settingsButtonDxfPart, true);
        }


        //обработка нажатия кнопки DXF
        private void MyButton_OnExecute(NameValueMap Context)
        {
            try
            {
                CreatDoc rule = new CreatDoc();
                rule.Rule();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обработки нажатия" + ex.Message);
            }
        }

        private void MyButton_OnExecuteDxfPart(NameValueMap Context)
        {
            try
            {
                CreatDXFOne rule1 = new CreatDXFOne();
                rule1.Rule1();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка обработки нажатия" + ex.Message);
            }
        }
    }
}