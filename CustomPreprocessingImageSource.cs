
using FCEngine;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Hello
{
    class CustomPreprocessingImageSource : IImageSource
    {
        public CustomPreprocessingImageSource(IEngine engine)
        {
            fileNames = new Queue();
            imageTools = engine.CreateImageProcessingTools();
            _languageRussian = engine.PredefinedLanguages.FindLanguage("Russian");
            _langEnglish = engine.PredefinedLanguages.FindLanguage("English");

        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(_languageRussian);
            Marshal.ReleaseComObject(_langEnglish);
            Marshal.ReleaseComObject(imageTools);
        }

        public void AddImageFile(string filePath)
        {
            fileNames.Enqueue(filePath);
            _nameSource = Path.GetFileNameWithoutExtension(filePath) + "_p_1";
        }


        /// <summary>
        /// Возвращает имя текущего изображения
        /// </summary>
        /// <returns>Строка - имя файла</returns>
        public string GetName() { return _nameSource; }
        public IFileAdapter GetNextImageFile()
        {
            // You could apply an external tool to a file and return the reference
            // to the resulting image file. This sample source does not use this feature
            return null;
        }
        public IImage GetNextImage()
        {
            while (true)
            {
                if (imageFile == null)
                {
                    if (fileNames.Count > 0)
                    {
                        //_nameSource = (string)fileNames.Dequeue();
                        imageFile = imageTools.OpenImageFile((string)fileNames.Dequeue());
                        nextPageIndex = 0;
                    }
                    else
                    {
                        return null;
                    }
                }
                Debug.Assert(imageFile != null);
                if (nextPageIndex < imageFile.PagesCount)
                {
                    IImage nextPage = imageFile.OpenImagePage(nextPageIndex++);
                    _nameSource = _nameSource.Remove(_nameSource.Length - 1) + (nextPageIndex + 1).ToString();
                    // корректировка разрешения
                    if (!Use600dpiRecognition)
                    {
                        if (_isResolutionChange)
                        {
                            if (nextPage.Resolution != _valueResolutionImage)
                                nextPage = imageTools.StretchImage(nextPage, nextPage.Width, nextPage.Height, _valueResolutionImage);
                        }
                        else
                        {
                            if (nextPage.Resolution == 96)
                                nextPage = imageTools.StretchImage(nextPage, nextPage.Width, nextPage.Height, 170);
                            else if (nextPage.Resolution < 150 || nextPage.Resolution > 600)
                                nextPage = imageTools.StretchImage(nextPage, nextPage.Width, nextPage.Height, 300);
                        }
                    }
                    else
                    {
                        if (nextPage.Resolution != 600)
                            nextPage = imageTools.StretchImage(nextPage, nextPage.Width, nextPage.Height, 600);
                    }

                    SaveInitialValues();
                    AdjustPreprocessingProperties(nextPage);

                    // дополнительная корректировка изображений                   
                    if (_condIsPhoto)
                    {
                        nextPage = PreprocessPhotoImg(nextPage);
                    }
                    else if (_isColorBack)
                    {
                        nextPage = PreprocessColorBackImg(nextPage);
                    }
                    else if (_isQualityImageBad)
                        nextPage = CorrectBadImages(nextPage);

                    // проверка ориентации страницы
                    nextPage = FindOrientation(_languageRussian, nextPage);
                    if (!_isOrientationChanged)
                        nextPage = FindOrientation(_langEnglish, nextPage);

                    // конвертация в черно-белое 
                    if (_binarizeImages)
                    {
                        if (!_condIsPhoto)
                            nextPage = imageTools.Binarize(nextPage, _modeBinarization, false, false);
                        //nextPage = imageTools.Binarize(nextPage, BinarizationModeEnum.BM_ByThreshold, false, false, CalculateThreshold(nextPage.Resolution));
                        else
                            nextPage = imageTools.Binarize(nextPage, BinarizationModeEnum.BM_Default, true, false); // SmoothTexture - false ???

                    }
                    else if (ConvertToGrey)
                    {
                        nextPage = imageTools.ConvertToGray(nextPage);
                    }

                    // выполнить очистку мусора на изображении
                    if (_isGarbageRemove)
                        nextPage = imageTools.RemoveGarbage(nextPage, _sizePixelGarbage);

#if debug

                    if (AppParameters.LoggingMode)
                        LogObjectPropirties();
#endif
                    RestoreInitialValues();
                    return nextPage;
                }
                else
                {
                    imageFile = null;
                    continue;
                }
            }
        }

        /// <summary>
        /// Разворачиваем изображение, если оно не горизонтально
        /// </summary>
        /// <param name="lang">Искомый язык</param>
        /// <param name="nextPage">Исходное изображение</param>
        /// <returns>Полученное изображение</returns>
        IImage FindOrientation(ILanguage lang, IImage nextPage)
        {
            // 
            RotationTypeEnum rotateTo = imageTools.DetectOrientationByText(nextPage, lang);

            if (rotateTo != RotationTypeEnum.RT_NoRotation)
            {
                if (rotateTo == RotationTypeEnum.RT_Counterclockwise)
                    rotateTo = RotationTypeEnum.RT_Clockwise;
                else if (rotateTo == RotationTypeEnum.RT_Clockwise)
                    rotateTo = RotationTypeEnum.RT_Counterclockwise;

                nextPage = imageTools.RotateImageByRotationType(nextPage, rotateTo);

                _isOrientationChanged = true;
            }

            return nextPage;
        }

        /// <summary>
        /// Предобработка изображения с цветным фоном для улучшения качества распознавания 
        /// </summary>
        /// <param name="nextPage">модифицируемое изображение</param>
        /// <returns>отредактированное изображение</returns>
        IImage PreprocessColorBackImg(IImage nextPage)
        {
            if (checkIsColorImage(nextPage))
            {
                nextPage = imageTools.PreprocessCameraImage(nextPage, PreprocessCameraImagePresetEnum.PCIP_General);
                //nextPage = imageTools.EqualizeBrightness(nextPage); // хорошо работает с гербовыми печатями и пересканами(?), бывает, что качество ухудшается на некоторых изображениях
                //nextPage = imageTools.RemoveMotionBlur(nextPage); // может полностью испортить изображение
                //nextPage = imageTools.CorrectImageGeometry(nextPage, true); // вроде бы бесполезно, но если второй параметр - true, ухудшается качество местами
                //nextPage = imageTools.RemoveNoise(nextPage, null);
                //nextPage = imageTools.FilterColor(nextPage, null, ColorToFilterEnum.CTF_Red, FilterColorModeEnum.FCM_Background);
                //nextPage = imageTools.SuppressColorObjects(nextPage, 122, 25);

                /* // хорошая альтернатива функции PreprocessCameraImage
                IFrequencyTransformParams transformParams = imageTools.CreateFrequencyTransformParams();
                transformParams.Type = FrequencyTransformTypeEnum.FTT_MultiplyByCoefficient;
                transformParams.NumberOfLevels = 10;
                transformParams.set_CoefficientAtLevel(7, 0.0);
                transformParams.set_CoefficientAtLevel(8, 0.0);
                transformParams.set_CoefficientAtLevel(9, 0.0);
                nextPage = imageTools.ApplyFrequencyTransform(nextPage, transformParams);*/


                //nextPage = ImproveCharactersBySize(nextPage);

            }

            if (checkIsColorOrGrayImage(nextPage))
                nextPage = imageTools.SmoothTexture(nextPage);

            return nextPage;
        }

        /// <summary>
        /// Предобработка изображения снятых на камеру для улучшения качества распознавания 
        /// </summary>
        /// <param name="nextPage">модифицируемое изображение</param>
        /// <returns>отредактированное изображение</returns>
        IImage PreprocessPhotoImg(IImage nextPage)
        {
            if (checkIsColorImage(nextPage))
                nextPage = imageTools.PreprocessCameraImage(nextPage, PreprocessCameraImagePresetEnum.PCIP_General);

            if (checkIsColorOrGrayImage(nextPage))
                nextPage = imageTools.SmoothTexture(nextPage);
            nextPage = imageTools.CorrectImageGeometry(nextPage, true);
            //nextPage = imageTools.RemoveMotionBlur(nextPage);
            //imageTools.CreateDenoiseFilterParamsByImage(nextPage, NoiseModelTypeEnum.NMT_Normal, 12); //???
            //imageTools.RemoveNoise();
            return nextPage;
        }

        bool checkIsColorImage(IImage image)
        {
            if (image.ImageColorType == ImageColorTypeEnum.ICT_Color)
            {
                return true;
            }
            else
                return false;
        }

        bool checkIsColorOrGrayImage(IImage image)
        {
            if (image.ImageColorType == ImageColorTypeEnum.ICT_Color || image.ImageColorType == ImageColorTypeEnum.ICT_Gray)
            {
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Извлекает фон изображения
        /// </summary>
        /// <param name="page">модифицируемое изображение</param>
        /// <returns>фоновое изображение</returns>
        IImage ExtractBackground(IImage page)
        {
            var transformParams = imageTools.CreateFrequencyTransformParams();
            transformParams.Type = FrequencyTransformTypeEnum.FTT_MultiplyByCoefficient;
            transformParams.NumberOfLevels = 10;
            for (int i = 0; i < transformParams.NumberOfLevels; i++)
            {
                transformParams.set_CoefficientAtLevel(i, 0.0);
            }

            page = imageTools.ApplyFrequencyTransform(page, transformParams);
            return page;
        }

        /// <summary>Корректирует яркость изображений</summary>
        /// <param name="page"></param>
        IImage EqualizingBrightness(IImage page) // NOTICE: есть встроенная функция
        {
            var transformParams = imageTools.CreateFrequencyTransformParams();
            transformParams.Type = FrequencyTransformTypeEnum.FTT_MultiplyByCoefficient;
            transformParams.NumberOfLevels = 10;
            for (int i = 0; i < transformParams.NumberOfLevels; i++)
            {
                transformParams.set_CoefficientAtLevel(i, 1.0);
            }
            for (int i = 7; i < transformParams.NumberOfLevels; i++)
            {
                transformParams.set_CoefficientAtLevel(i, 0.0);
            }
            page = imageTools.ApplyFrequencyTransform(page, transformParams);
            return page;
        }

        /// <summary>
        /// Повышает контрастность изображения для выделения текста
        /// </summary>
        /// <param name="page">модифицируемое изображение</param>
        /// <returns>отредактированное изображение</returns>
        IImage ImproveCharactersContrast(IImage page)
        {
            // !!! принудительно отключать конвертацию в ч/б

            var transformParams = imageTools.CreateFrequencyTransformParams();
            transformParams.Type = FrequencyTransformTypeEnum.FTT_AutocontrastByCoefficient;
            transformParams.NumberOfLevels = 5;
            transformParams.set_CoefficientAtLevel(2, 1.0);
            transformParams.set_CoefficientAtLevel(3, 1.0);
            transformParams.set_CoefficientAtLevel(4, 1.0);
            page = imageTools.ApplyFrequencyTransform(page, transformParams);
            return page;
        }

        /// <summary>Улучшает распознавание символово определенного размера</summary>
        /// <param name="page"></param>
        IImage ImproveCharactersBySize(IImage page)
        {
            // !!! принудительно отключать конвертацию в ч/б

            int sizeChars = 15;
            var transformParams = imageTools.CreateFrequencyTransformParamsByCharSize(sizeChars);
            page = imageTools.ApplyFrequencyTransform(page, transformParams);
            return page;
        }

        /// <summary>Улучшает эффективность распознавание изображений низкого качества</summary>
        /// <param name="pImage"></param>
        IImage CorrectBadImages(IImage pImage)
        {
            // Серый режим в большинстве случаев работает лучше ч/б


            pImage = ImproveCharactersBySize(pImage);
            pImage = imageTools.CorrectImageGeometry(pImage, false);
            pImage = imageTools.EqualizeBrightness(pImage);

            //pImage = imageTools.SmoothTexture(pImage);

            return pImage;
        }

        /// <summary>Если многостраничный документ имеет разную длину страниц, то он приводятся к длине первой страницы</summary>
        /// <param name="pImage"></param>
        IImage FixUnequalWidth(IImage pImage)
        {
            return pImage;
        }

        int CalculateThreshold(int pResolution)
        {
            int valueThreshold = 200;

            if (pResolution < 151)
                valueThreshold = 90;
            else if (pResolution < 251)
                valueThreshold = 60;
            else if (pResolution < 351)
                valueThreshold = 30;
            else if (pResolution < 451)
                valueThreshold = 20;
            else if (pResolution < 601)
                valueThreshold = 15;

            return valueThreshold;
        }

        void SaveImageToDisk(IImage image, string pathToSave)
        {
            image.WriteToFile(pathToSave, ImageFileFormatEnum.IFF_Tif, null);
        }

        /// <summary>Выполняет настройку параметров предобработки изображений</summary>
        /// <param name="imageAdjasting">Изображение для тестирования</param>
        void AdjustPreprocessingProperties(IImage imageAdjasting)
        {
            if (!_useAutomaticPreprocessing)
            {
                if (IsColorBackImage)
                {
                    _binarizeImages = false;
                    _condIsPhoto = false;
                }
                else if (IsPhotoImage)
                {
                    _isColorBack = false;
                }
            }

            // проверка, что изображение не ч/б
            if (!checkIsColorOrGrayImage(imageAdjasting))
            {
                _condIsPhoto = false;
                _isColorBack = false;
                _binarizeImages = false;

            }
            else
            {
                if (!_binarizeImages)
                    _isGarbageRemove = false;
            }

            if (AppParameters._isSuitableForOcr)
            {
                if (String.IsNullOrEmpty(AppParameters._warnColorTypeBad))
                    if (checkIsColorImage(imageAdjasting))
                        AppParameters._warnColorTypeBad = "Некоторые изображения сканированы в цветном режиме, поменяйте на оттенки серого или черно-белый";

                if (String.IsNullOrEmpty(AppParameters._warnResolutionImageBad))
                    if (imageAdjasting.Resolution < 300)
                        AppParameters._warnResolutionImageBad = "Оптимальное разрешение сканера для распознавания должно быть 300 - 400 dpi, если текст очень маленький - 600 dpi";


                if (!String.IsNullOrEmpty(AppParameters._warnResolutionImageBad) && !String.IsNullOrEmpty(AppParameters._warnColorTypeBad))
                    AppParameters._isSuitableForOcr = false;
            }
        }

        void LogObjectProperties()
        {
            string createText = "===Class properties===" + Environment.NewLine;
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(this))
            {
                string name = descriptor.Name;
                object value = descriptor.GetValue(this);
                createText += name + "=" + value.ToString() + Environment.NewLine;
            }
            createText += "===End class properties===" + Environment.NewLine;

            try
            {
                File.AppendAllText(AppParameters.PathFileLog, createText);
            }
            catch (Exception ex)
            {
                string textError = "LogObjectPropirties: Не удалось записать информацию в файл лога" + Environment.NewLine + ex.Message;
                AppParameters.TextError.Add(textError);
                throw;
            }
        }

        /// <summary>Сохраняет заданные параметры предобработки во временные переменные</summary>
        void SaveInitialValues()
        {
            _condIsPhotoTmp = _condIsPhoto;
            _binarizeImagesTmp = _binarizeImages;
            _doGreyConvertTmp = _doGreyConvert;
            _isColorBackTmp = _isColorBack;
            _isGarbageRemoveTmp = _isGarbageRemove;
        }

        /// <summary>Восстанавливает значения параметров предобработки изображений</summary>
        void RestoreInitialValues()
        {
            _condIsPhoto = _condIsPhotoTmp;
            _binarizeImages = _binarizeImagesTmp;
            _isColorBack = _isColorBackTmp;
            _doGreyConvert = _doGreyConvertTmp;
            _isGarbageRemove = _isGarbageRemoveTmp;
        }

        #region IMPLEMENTATION

        IImageProcessingTools imageTools;

        Queue fileNames;
        IImageFile imageFile;
        ILanguage _languageRussian;
        ILanguage _langEnglish;
        int nextPageIndex;
        bool _condIsPhoto = false;
        bool _binarizeImages = true;
        private bool _doGreyConvert = false;
        bool _isColorBack = false;
        bool _is600dpi = false;
        bool _isGarbageRemove = false;
        int _sizePixelGarbage = 0;
        bool _useAutomaticPreprocessing = false;

        bool _isResolutionChange = false;
        int _valueResolutionImage = 0;

        bool _isOrientationChanged = false;
        bool _isQualityImageBad = false;

        // временные значения для хранения
        bool _condIsPhotoTmp = false;
        bool _binarizeImagesTmp = true;
        private bool _doGreyConvertTmp = false;
        bool _isColorBackTmp = false;
        bool _isGarbageRemoveTmp = false;

        // возможно по умолчанию стоит использовать BM_ByThreshold, тк символы распознаются точнее и лучше исправляется перекос
        // но нужно написать функцию оптимизации параметра Thrashold для функции Binarize
        BinarizationModeEnum _modeBinarization = BinarizationModeEnum.BM_Fast;

        public BinarizationModeEnum ModeBinarization { get { return _modeBinarization; } set { _modeBinarization = value; } }

        public bool UseBinarization { get { return _binarizeImages; } set { _binarizeImages = value; _doGreyConvert = !value; } }

        public bool ConvertToGrey
        {
            get { return _doGreyConvert; }
            set
            {
                _doGreyConvert = value;
                _binarizeImages = !value;
            }
        }

        public bool UseModeAutomatic { get { return _useAutomaticPreprocessing; } set { _useAutomaticPreprocessing = value; } }

        public bool IsColorBackImage { get { return _isColorBack; } set { _isColorBack = value; if (value) _condIsPhoto = false; } }

        public bool IsPhotoImage { get { return _condIsPhoto; } set { _condIsPhoto = value; if (value) _isColorBack = false; } }

        public bool Use600dpiRecognition { get { return _is600dpi; } set { _is600dpi = value; } }

        public bool UseCleanUpImage { get { return _isGarbageRemove; } set { _isGarbageRemove = value; } }
        public int SetSizeGarbageRemove { get { return _sizePixelGarbage; } set { if (value >= 0) _sizePixelGarbage = value; } }

        public int SetResolutionImage
        {
            get { return _valueResolutionImage; }
            set
            {
                if (value > 0) { _isResolutionChange = true; if (value < 98) value = 98; _valueResolutionImage = value; }
                else { _isResolutionChange = false; _valueResolutionImage = 0; }
            }
        }


        string _nameSource = "PreprocessedImage";

        #endregion
    };

    /// <summary>Доступные настройки предобработки изображений для распознавания в облаке</summary>
    public static class ImagePreprocessingSettings
    {
        /// <summary>
        ///   <para>Задает режим очистки изображений от мусорных элементов</para>
        /// </summary>
        /// <value>минус 1 - очистка отключена. 0 - включить автоматическую очистку, число больше 0 - очистка элементов до заданного размера</value>
        /// <remarks>Данный режим не работает, если изображение не ч/б и предобработка изображений отключена</remarks>
        public static int UseGarbageRemove { get; set; } = -1;

        /// <summary>Задать режим предобработки изображений</summary>
        /// <value>0 - Предобработка изображений отключена, 1- Стандартная предобработка, 2 - Альтернативная предобработка, 4 - Автоматический режим</value>
        public static int valueBinarizationMethod { get; set; } = 1;

        /// <summary>Принудительная конвертация изображений в разрешение 600 dpi</summary>
        /// <value>Истина - включить</value>
        /// <remarks>Данный режим предобработки нужен для повышения точности распознавания изображений с мелким текстом</remarks>
        public static bool is600dpi { get; set; } = false;

        /// <summary>Задать режим обработки документов с цветным фоном</summary>
        /// <value>Истина - включить</value>
        /// <remarks>Примеры документов: Паспорт, СНИЛС, ИНН</remarks>
        public static bool isColorBackImage { get; set; } = false;

        /// <summary>Задать режим предобработки изображений полученных с фото камеры</summary>
        /// <value>Истина - включить</value>
        public static bool isPhoto { get; set; } = false;

        /// <summary>Принудительная конвертация изображений к заданному разрешению</summary>
        /// <value>0 - не конвертировать, значение больше 97 - новое разрешение</value>
        public static int SetResolutionImage { get; set; } = 0;

    };
}
