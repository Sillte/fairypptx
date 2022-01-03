from ._constants import MSO    # NOQA
from ._constants import MSPPT  # NOQA

class constants(MSO.constants, MSPPT.constants):
    pass

# ### 
# We want to use `constants` as `from fairypptx.constants import msoTrue`. 
# 
# It is all right to execute below code, 
# 
# ```
# d = globals()
# for key in dir(constants):
#    d[key] = getattr(constants, key)
# ```
# However, we cannot lookup the defintion of each constant.  
# To counter-act against this problem, 
# generated the variables naively below. 


"""To generate this code,
[pywin32](https://github.com/mhammond/pywin32) is used.
Thank you very much !!
"""

BackstageGroupStyleError = 2  # Enum: BackstageGroupStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.backstagegroupstyle
BackstageGroupStyleNormal = 0  # Enum: BackstageGroupStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.backstagegroupstyle
BackstageGroupStyleWarning = 1  # Enum: BackstageGroupStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.backstagegroupstyle
certdetAvailable = 0  # Enum: CertificateDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificatedetail
certdetExpirationDate = 3  # Enum: CertificateDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificatedetail
certdetIssuer = 2  # Enum: CertificateDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificatedetail
certdetSubject = 1  # Enum: CertificateDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificatedetail
certdetThumbprint = 4  # Enum: CertificateDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificatedetail
certverresError = 0  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresExpired = 5  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresInvalid = 4  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresRevoked = 6  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresUntrusted = 7  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresUnverified = 2  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresValid = 3  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
certverresVerifying = 1  # Enum: CertificateVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.certificateverificationresults
contverresError = 0  # Enum: ContentVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.contentverificationresults
contverresModified = 4  # Enum: ContentVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.contentverificationresults
contverresUnverified = 2  # Enum: ContentVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.contentverificationresults
contverresValid = 3  # Enum: ContentVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.contentverificationresults
contverresVerifying = 1  # Enum: ContentVerificationResults, https://docs.microsoft.com/ja-jp/office/vba/api/office.contentverificationresults
offPropertyTypeBoolean = 2  # Enum: DocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.docproperties
offPropertyTypeDate = 3  # Enum: DocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.docproperties
offPropertyTypeFloat = 5  # Enum: DocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.docproperties
offPropertyTypeNumber = 1  # Enum: DocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.docproperties
offPropertyTypeString = 4  # Enum: DocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.docproperties
cipherModeCBC = 1  # Enum: EncryptionCipherMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionciphermode
cipherModeECB = 0  # Enum: EncryptionCipherMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionciphermode
encprovdetAlgorithm = 1  # Enum: EncryptionProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionproviderdetail
encprovdetBlockCipher = 2  # Enum: EncryptionProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionproviderdetail
encprovdetCipherBlockSize = 3  # Enum: EncryptionProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionproviderdetail
encprovdetCipherMode = 4  # Enum: EncryptionProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionproviderdetail
encprovdetUrl = 0  # Enum: EncryptionProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.encryptionproviderdetail
mfHTML = 2  # Enum: MailFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.mailformat
mfPlainText = 1  # Enum: MailFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.mailformat
mfRTF = 3  # Enum: MailFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.mailformat
msoAlertButtonAbortRetryIgnore = 2  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertButtonOK = 0  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertButtonOKCancel = 1  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertButtonRetryCancel = 5  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertButtonYesAllNoCancel = 6  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertButtonYesNo = 4  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertButtonYesNoCancel = 3  # Enum: MsoAlertButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertbuttontype
msoAlertCancelDefault = (
    -1
)  # Enum: MsoAlertCancelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertcanceltype
msoAlertCancelFifth = 4  # Enum: MsoAlertCancelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertcanceltype
msoAlertCancelFirst = 0  # Enum: MsoAlertCancelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertcanceltype
msoAlertCancelFourth = 3  # Enum: MsoAlertCancelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertcanceltype
msoAlertCancelSecond = 1  # Enum: MsoAlertCancelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertcanceltype
msoAlertCancelThird = 2  # Enum: MsoAlertCancelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertcanceltype
msoAlertDefaultFifth = 4  # Enum: MsoAlertDefaultType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertdefaulttype
msoAlertDefaultFirst = 0  # Enum: MsoAlertDefaultType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertdefaulttype
msoAlertDefaultFourth = 3  # Enum: MsoAlertDefaultType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertdefaulttype
msoAlertDefaultSecond = 1  # Enum: MsoAlertDefaultType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertdefaulttype
msoAlertDefaultThird = 2  # Enum: MsoAlertDefaultType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalertdefaulttype
msoAlertIconCritical = 1  # Enum: MsoAlertIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalerticontype
msoAlertIconInfo = 4  # Enum: MsoAlertIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalerticontype
msoAlertIconNoIcon = 0  # Enum: MsoAlertIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalerticontype
msoAlertIconQuery = 2  # Enum: MsoAlertIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalerticontype
msoAlertIconWarning = 3  # Enum: MsoAlertIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoalerticontype
msoAlignBottoms = 5  # Enum: MsoAlignCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoaligncmd
msoAlignCenters = 1  # Enum: MsoAlignCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoaligncmd
msoAlignLefts = 0  # Enum: MsoAlignCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoaligncmd
msoAlignMiddles = 4  # Enum: MsoAlignCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoaligncmd
msoAlignRights = 2  # Enum: MsoAlignCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoaligncmd
msoAlignTops = 3  # Enum: MsoAlignCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoaligncmd
msoAnimationAppear = 32  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationBeginSpeaking = 4  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationCharacterSuccessMajor = 6  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationCheckingSomething = 103  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationDisappear = 31  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationEmptyTrash = 116  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGestureDown = 113  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGestureLeft = 114  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGestureRight = 19  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGestureUp = 115  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGetArtsy = 100  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGetAttentionMajor = 11  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGetAttentionMinor = 12  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGetTechy = 101  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGetWizardy = 102  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGoodbye = 3  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationGreeting = 2  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationIdle = 1  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationListensToComputer = 26  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookDown = 104  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookDownLeft = 105  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookDownRight = 106  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookLeft = 107  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookRight = 108  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookUp = 109  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookUpLeft = 110  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationLookUpRight = 111  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationPrinting = 18  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationRestPose = 5  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationSaving = 112  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationSearching = 13  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationSendingMail = 25  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationThinking = 24  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationWorkingAtSomething = 23  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoAnimationWritingNotingSomething = 22  # Enum: MsoAnimationType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoanimationtype
msoLanguageIDExeMode = 4  # Enum: MsoAppLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoapplanguageid
msoLanguageIDHelp = 3  # Enum: MsoAppLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoapplanguageid
msoLanguageIDInstall = 1  # Enum: MsoAppLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoapplanguageid
msoLanguageIDUI = 2  # Enum: MsoAppLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoapplanguageid
msoLanguageIDUIPrevious = 5  # Enum: MsoAppLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoapplanguageid
msoArrowheadLengthMedium = 2  # Enum: MsoArrowheadLength, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadlength
msoArrowheadLengthMixed = (
    -2
)  # Enum: MsoArrowheadLength, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadlength
msoArrowheadLong = 3  # Enum: MsoArrowheadLength, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadlength
msoArrowheadShort = 1  # Enum: MsoArrowheadLength, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadlength
msoArrowheadDiamond = 5  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadNone = 1  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadOpen = 3  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadOval = 6  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadStealth = 4  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadStyleMixed = (
    -2
)  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadTriangle = 2  # Enum: MsoArrowheadStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadstyle
msoArrowheadNarrow = 1  # Enum: MsoArrowheadWidth, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadwidth
msoArrowheadWide = 3  # Enum: MsoArrowheadWidth, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadwidth
msoArrowheadWidthMedium = 2  # Enum: MsoArrowheadWidth, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadwidth
msoArrowheadWidthMixed = (
    -2
)  # Enum: MsoArrowheadWidth, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoarrowheadwidth
msoShape10pointStar = 149  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape12pointStar = 150  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape16pointStar = 94  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape24pointStar = 95  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape32pointStar = 96  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape4pointStar = 91  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape5pointStar = 92  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape6pointStar = 147  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape7pointStar = 148  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShape8pointStar = 93  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonBackorPrevious = 129  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonBeginning = 131  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonCustom = 125  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonDocument = 134  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonEnd = 132  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonForwardorNext = 130  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonHelp = 127  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonHome = 126  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonInformation = 128  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonMovie = 136  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonReturn = 133  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeActionButtonSound = 135  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeArc = 25  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeBalloon = 137  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeBentArrow = 41  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeBentUpArrow = 44  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeBevel = 15  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeBlockArc = 20  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCan = 13  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeChartPlus = 182  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeChartStar = 181  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeChartX = 180  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeChevron = 52  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeChord = 161  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCircularArrow = 60  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCloud = 179  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCloudCallout = 108  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCorner = 162  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCornerTabs = 169  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCross = 11  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCube = 14  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCurvedDownArrow = 48  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCurvedDownRibbon = 100  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCurvedLeftArrow = 46  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCurvedRightArrow = 45  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCurvedUpArrow = 47  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeCurvedUpRibbon = 99  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDecagon = 144  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDiagonalStripe = 141  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDiamond = 4  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDodecagon = 146  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDonut = 18  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDoubleBrace = 27  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDoubleBracket = 26  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDoubleWave = 104  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDownArrow = 36  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDownArrowCallout = 56  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeDownRibbon = 98  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeExplosion1 = 89  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeExplosion2 = 90  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartAlternateProcess = 62  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartCard = 75  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartCollate = 79  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartConnector = 73  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartData = 64  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartDecision = 63  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartDelay = 84  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartDirectAccessStorage = 87  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartDisplay = 88  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartDocument = 67  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartExtract = 81  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartInternalStorage = 66  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartMagneticDisk = 86  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartManualInput = 71  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartManualOperation = 72  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartMerge = 82  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartMultidocument = 68  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartOfflineStorage = 139  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartOffpageConnector = 74  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartOr = 78  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartPredefinedProcess = 65  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartPreparation = 70  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartProcess = 61  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartPunchedTape = 76  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartSequentialAccessStorage = 85  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartSort = 80  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartStoredData = 83  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartSummingJunction = 77  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFlowchartTerminator = 69  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFoldedCorner = 16  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFrame = 158  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeFunnel = 174  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeGear6 = 172  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeGear9 = 173  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeHalfFrame = 159  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeHeart = 21  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeHeptagon = 145  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeHexagon = 10  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeHorizontalScroll = 102  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeIsoscelesTriangle = 7  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftArrow = 34  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftArrowCallout = 54  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftBrace = 31  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftBracket = 29  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftCircularArrow = 176  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftRightArrow = 37  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftRightArrowCallout = 57  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftRightCircularArrow = 177  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftRightRibbon = 140  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftRightUpArrow = 40  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLeftUpArrow = 43  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLightningBolt = 22  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout1 = 109  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout1AccentBar = 113  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout1BorderandAccentBar = 121  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout1NoBorder = 117  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout2 = 110  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout2AccentBar = 114  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout2BorderandAccentBar = 122  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout2NoBorder = 118  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout3 = 111  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout3AccentBar = 115  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout3BorderandAccentBar = 123  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout3NoBorder = 119  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout4 = 112  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout4AccentBar = 116  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout4BorderandAccentBar = 124  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineCallout4NoBorder = 120  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeLineInverse = 183  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMathDivide = 166  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMathEqual = 167  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMathMinus = 164  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMathMultiply = 165  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMathNotEqual = 168  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMathPlus = 163  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMixed = (
    -2
)  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeMoon = 24  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeNoSymbol = 19  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeNonIsoscelesTrapezoid = 143  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeNotPrimitive = 138  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeNotchedRightArrow = 50  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeOctagon = 6  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeOval = 9  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeOvalCallout = 107  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeParallelogram = 2  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapePentagon = 51  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapePie = 142  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapePieWedge = 175  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapePlaque = 28  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapePlaqueTabs = 171  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeQuadArrow = 39  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeQuadArrowCallout = 59  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRectangle = 1  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRectangularCallout = 105  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRegularPentagon = 12  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRightArrow = 33  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRightArrowCallout = 53  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRightBrace = 32  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRightBracket = 30  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRightTriangle = 8  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRound1Rectangle = 151  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRound2DiagRectangle = 153  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRound2SameRectangle = 152  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRoundedRectangle = 5  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeRoundedRectangularCallout = 106  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSmileyFace = 17  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSnip1Rectangle = 155  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSnip2DiagRectangle = 157  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSnip2SameRectangle = 156  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSnipRoundRectangle = 154  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSquareTabs = 170  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeStripedRightArrow = 49  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSun = 23  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeSwooshArrow = 178  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeTear = 160  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeTrapezoid = 3  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeUTurnArrow = 42  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeUpArrow = 35  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeUpArrowCallout = 55  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeUpDownArrow = 38  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeUpDownArrowCallout = 58  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeUpRibbon = 97  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeVerticalScroll = 101  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoShapeWave = 103  # Enum: MsoAutoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautoshapetype
msoAutoSizeMixed = (
    -2
)  # Enum: MsoAutoSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautosize
msoAutoSizeNone = 0  # Enum: MsoAutoSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautosize
msoAutoSizeShapeToFitText = 1  # Enum: MsoAutoSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautosize
msoAutoSizeTextToFitShape = 2  # Enum: MsoAutoSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautosize
msoAutomationSecurityByUI = 2  # Enum: MsoAutomationSecurity, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautomationsecurity
msoAutomationSecurityForceDisable = 3  # Enum: MsoAutomationSecurity, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautomationsecurity
msoAutomationSecurityLow = 1  # Enum: MsoAutomationSecurity, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoautomationsecurity
msoBackgroundStyleMixed = (
    -2
)  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStyleNotAPreset = 0  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset1 = 1  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset10 = 10  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset11 = 11  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset12 = 12  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset2 = 2  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset3 = 3  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset4 = 4  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset5 = 5  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset6 = 6  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset7 = 7  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset8 = 8  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBackgroundStylePreset9 = 9  # Enum: MsoBackgroundStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobackgroundstyleindex
msoBalloonButtonAbort = (
    -8
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonBack = (
    -5
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonCancel = (
    -2
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonClose = (
    -12
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonIgnore = (
    -9
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonNext = (
    -6
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonNo = (
    -4
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonNull = 0  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonOK = (
    -1
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonOptions = (
    -14
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonRetry = (
    -7
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonSearch = (
    -10
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonSnooze = (
    -11
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonTips = (
    -13
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonYes = (
    -3
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonButtonYesToAll = (
    -15
)  # Enum: MsoBalloonButtonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonbuttontype
msoBalloonErrorBadCharacter = 8  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorBadPictureRef = 4  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorBadReference = 5  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorButtonModeless = 7  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorButtonlessModal = 6  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorCOMFailure = 9  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorCharNotTopmostForModal = 10  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorNone = 0  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorOther = 1  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorOutOfMemory = 3  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorTooBig = 2  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonErrorTooManyControls = 11  # Enum: MsoBalloonErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoonerrortype
msoBalloonTypeBullets = 1  # Enum: MsoBalloonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoontype
msoBalloonTypeButtons = 0  # Enum: MsoBalloonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoontype
msoBalloonTypeNumbers = 2  # Enum: MsoBalloonType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoballoontype
msoBarBottom = 3  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarFloating = 4  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarLeft = 0  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarMenuBar = 6  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarPopup = 5  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarRight = 2  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarTop = 1  # Enum: MsoBarPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarposition
msoBarNoChangeDock = 16  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoChangeVisible = 8  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoCustomize = 1  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoHorizontalDock = 64  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoMove = 4  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoProtection = 0  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoResize = 2  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarNoVerticalDock = 32  # Enum: MsoBarProtection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarprotection
msoBarRowFirst = 0  # Enum: MsoBarRow, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarrow
msoBarRowLast = (
    -1
)  # Enum: MsoBarRow, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobarrow
msoBarTypeMenuBar = 1  # Enum: MsoBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobartype
msoBarTypeNormal = 0  # Enum: MsoBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobartype
msoBarTypePopup = 2  # Enum: MsoBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobartype
msoBaselineAlignAuto = 5  # Enum: MsoBaselineAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobaselinealignment
msoBaselineAlignBaseline = 1  # Enum: MsoBaselineAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobaselinealignment
msoBaselineAlignCenter = 3  # Enum: MsoBaselineAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobaselinealignment
msoBaselineAlignFarEast50 = 4  # Enum: MsoBaselineAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobaselinealignment
msoBaselineAlignMixed = (
    -2
)  # Enum: MsoBaselineAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobaselinealignment
msoBaselineAlignTop = 2  # Enum: MsoBaselineAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobaselinealignment
msoBevelAngle = 6  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelArtDeco = 13  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelCircle = 3  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelConvex = 8  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelCoolSlant = 9  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelCross = 5  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelDivot = 10  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelHardEdge = 12  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelNone = 1  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelRelaxedInset = 2  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelRiblet = 11  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelSlope = 4  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelSoftRound = 7  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBevelTypeMixed = (
    -2
)  # Enum: MsoBevelType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobeveltype
msoBlackWhiteAutomatic = 1  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteBlack = 8  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteBlackTextAndLine = 6  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteDontShow = 10  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteGrayOutline = 5  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteGrayScale = 2  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteHighContrast = 7  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteInverseGrayScale = 4  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteLightGrayScale = 3  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteMixed = (
    -2
)  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlackWhiteWhite = 9  # Enum: MsoBlackWhiteMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblackwhitemode
msoBlogMultipleCategories = 2  # Enum: MsoBlogCategorySupport, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblogcategorysupport
msoBlogNoCategories = 0  # Enum: MsoBlogCategorySupport, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblogcategorysupport
msoBlogOneCategory = 1  # Enum: MsoBlogCategorySupport, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblogcategorysupport
msoblogImageTypeGIF = 2  # Enum: MsoBlogImageType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblogimagetype
msoblogImageTypeJPEG = 1  # Enum: MsoBlogImageType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblogimagetype
msoblogImageTypePNG = 3  # Enum: MsoBlogImageType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoblogimagetype
BroadcastCapFileSizeLimited = 1  # Enum: MsoBroadcastCapabilities, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobroadcastcapabilities
BroadcastCapSupportsMeetingNotes = 2  # Enum: MsoBroadcastCapabilities, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobroadcastcapabilities
BroadcastCapSupportsUpdateDoc = 4  # Enum: MsoBroadcastCapabilities, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobroadcastcapabilities
BroadcastPaused = 2  # Enum: MsoBroadcastState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobroadcaststate
BroadcastStarted = 1  # Enum: MsoBroadcastState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobroadcaststate
NoBroadcast = 0  # Enum: MsoBroadcastState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobroadcaststate
msoBulletMixed = (
    -2
)  # Enum: MsoBulletType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobullettype
msoBulletNone = 0  # Enum: MsoBulletType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobullettype
msoBulletNumbered = 2  # Enum: MsoBulletType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobullettype
msoBulletPicture = 3  # Enum: MsoBulletType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobullettype
msoBulletUnnumbered = 1  # Enum: MsoBulletType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobullettype
msoButtonSetAbortRetryIgnore = 10  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetBackClose = 6  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetBackNextClose = 8  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetBackNextSnooze = 12  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetCancel = 2  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetNextClose = 7  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetNone = 0  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetOK = 1  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetOkCancel = 3  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetRetryCancel = 9  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetSearchClose = 11  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetTipsOptionsClose = 13  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetYesAllNoCancel = 14  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetYesNo = 4  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonSetYesNoCancel = 5  # Enum: MsoButtonSetType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonsettype
msoButtonDown = (
    -1
)  # Enum: MsoButtonState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstate
msoButtonMixed = 2  # Enum: MsoButtonState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstate
msoButtonUp = 0  # Enum: MsoButtonState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstate
msoButtonAutomatic = 0  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonCaption = 2  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonIcon = 1  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonIconAndCaption = 3  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonIconAndCaptionBelow = 11  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonIconAndWrapCaption = 7  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonIconAndWrapCaptionBelow = 15  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonWrapCaption = 14  # Enum: MsoButtonStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle
msoButtonTextBelow = 8  # Enum: MsoButtonStyleHidden, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstylehidden
msoButtonWrapText = 4  # Enum: MsoButtonStyleHidden, https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstylehidden
msoCTPDockPositionBottom = 3  # Enum: MsoCTPDockPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockposition
msoCTPDockPositionFloating = 4  # Enum: MsoCTPDockPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockposition
msoCTPDockPositionLeft = 0  # Enum: MsoCTPDockPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockposition
msoCTPDockPositionRight = 2  # Enum: MsoCTPDockPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockposition
msoCTPDockPositionTop = 1  # Enum: MsoCTPDockPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockposition
msoCTPDockPositionRestrictNoChange = 1  # Enum: MsoCTPDockPositionRestrict, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockpositionrestrict
msoCTPDockPositionRestrictNoHorizontal = 2  # Enum: MsoCTPDockPositionRestrict, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockpositionrestrict
msoCTPDockPositionRestrictNoVertical = 3  # Enum: MsoCTPDockPositionRestrict, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockpositionrestrict
msoCTPDockPositionRestrictNone = 0  # Enum: MsoCTPDockPositionRestrict, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoctpdockpositionrestrict
msoCalloutAngle30 = 2  # Enum: MsoCalloutAngleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutangletype
msoCalloutAngle45 = 3  # Enum: MsoCalloutAngleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutangletype
msoCalloutAngle60 = 4  # Enum: MsoCalloutAngleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutangletype
msoCalloutAngle90 = 5  # Enum: MsoCalloutAngleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutangletype
msoCalloutAngleAutomatic = 1  # Enum: MsoCalloutAngleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutangletype
msoCalloutAngleMixed = (
    -2
)  # Enum: MsoCalloutAngleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutangletype
msoCalloutDropBottom = 4  # Enum: MsoCalloutDropType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutdroptype
msoCalloutDropCenter = 3  # Enum: MsoCalloutDropType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutdroptype
msoCalloutDropCustom = 1  # Enum: MsoCalloutDropType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutdroptype
msoCalloutDropMixed = (
    -2
)  # Enum: MsoCalloutDropType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutdroptype
msoCalloutDropTop = 2  # Enum: MsoCalloutDropType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocalloutdroptype
msoCalloutFour = 4  # Enum: MsoCalloutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocallouttype
msoCalloutMixed = (
    -2
)  # Enum: MsoCalloutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocallouttype
msoCalloutOne = 1  # Enum: MsoCalloutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocallouttype
msoCalloutThree = 3  # Enum: MsoCalloutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocallouttype
msoCalloutTwo = 2  # Enum: MsoCalloutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocallouttype
msoCharacterSetArabic = 1  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetCyrillic = 2  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetEnglishWesternEuropeanOtherLatinScript = 3  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetGreek = 4  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetHebrew = 5  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetJapanese = 6  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetKorean = 7  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetMultilingualUnicode = 8  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetSimplifiedChinese = 9  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetThai = 10  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetTraditionalChinese = 11  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoCharacterSetVietnamese = 12  # Enum: MsoCharacterSet, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocharacterset
msoElementChartFloorNone = 1200  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementChartFloorShow = 1201  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementChartTitleAboveChart = 2  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementChartTitleCenteredOverlay = 1  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementChartTitleNone = 0  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementChartWallNone = 1100  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementChartWallShow = 1101  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelBestFit = 210  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelBottom = 209  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelCallout = 211  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelCenter = 202  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelInsideBase = 204  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelInsideEnd = 203  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelLeft = 206  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelNone = 200  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelOutSideEnd = 205  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelRight = 207  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelShow = 201  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataLabelTop = 208  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataTableNone = 500  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataTableShow = 501  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementDataTableWithLegendKeys = 502  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementErrorBarNone = 700  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementErrorBarPercentage = 702  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementErrorBarStandardDeviation = 703  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementErrorBarStandardError = 701  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendBottom = 104  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendLeft = 103  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendLeftOverlay = 106  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendNone = 100  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendRight = 101  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendRightOverlay = 105  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLegendTop = 102  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLineDropHiLoLine = 804  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLineDropLine = 801  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLineHiLoLine = 802  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLineNone = 800  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementLineSeriesLine = 803  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPlotAreaNone = 1000  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPlotAreaShow = 1001  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisBillions = 374  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisLogScale = 375  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisMillions = 373  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisNone = 348  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisReverse = 351  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisShow = 349  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisThousands = 372  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisTitleAdjacentToAxis = 301  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisTitleBelowAxis = 302  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisTitleHorizontal = 305  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisTitleNone = 300  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisTitleRotated = 303  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisTitleVertical = 304  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryAxisWithoutLabels = 350  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryGridLinesMajor = 334  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryGridLinesMinor = 333  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryGridLinesMinorMajor = 335  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryCategoryGridLinesNone = 332  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisBillions = 356  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisLogScale = 357  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisMillions = 355  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisNone = 352  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisShow = 353  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisThousands = 354  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisTitleAdjacentToAxis = 306  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisTitleBelowAxis = 308  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisTitleHorizontal = 311  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisTitleNone = 306  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisTitleRotated = 309  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueAxisTitleVertical = 310  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueGridLinesMajor = 330  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueGridLinesMinor = 329  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueGridLinesMinorMajor = 331  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementPrimaryValueGridLinesNone = 328  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisBillions = 378  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisLogScale = 379  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisMillions = 377  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisNone = 358  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisReverse = 361  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisShow = 359  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisThousands = 376  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisTitleAdjacentToAxis = 313  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisTitleBelowAxis = 314  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisTitleHorizontal = 317  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisTitleNone = 312  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisTitleRotated = 315  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisTitleVertical = 316  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryAxisWithoutLabels = 360  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryGridLinesMajor = 342  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryGridLinesMinor = 341  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryGridLinesMinorMajor = 343  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryCategoryGridLinesNone = 340  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisBillions = 366  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisLogScale = 367  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisMillions = 365  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisNone = 362  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisShow = 363  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisThousands = 364  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisTitleAdjacentToAxis = 319  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisTitleBelowAxis = 320  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisTitleHorizontal = 323  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisTitleNone = 318  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisTitleRotated = 321  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueAxisTitleVertical = 322  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueGridLinesMajor = 338  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueGridLinesMinor = 337  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueGridLinesMinorMajor = 339  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSecondaryValueGridLinesNone = 336  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisGridLinesMajor = 346  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisGridLinesMinor = 345  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisGridLinesMinorMajor = 347  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisGridLinesNone = 344  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisNone = 368  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisReverse = 371  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisShow = 369  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisTitleHorizontal = 327  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisTitleNone = 324  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisTitleRotated = 325  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisTitleVertical = 326  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementSeriesAxisWithoutLabeling = 370  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementTrendlineAddExponential = 602  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementTrendlineAddLinear = 601  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementTrendlineAddLinearForecast = 603  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementTrendlineAddTwoPeriodMovingAverage = 604  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementTrendlineNone = 600  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementUpDownBarsNone = 900  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoElementUpDownBarsShow = 901  # Enum: MsoChartElementType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartelementtype
msoChartFieldBubbleSize = 1  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoChartFieldCategoryName = 2  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoChartFieldFormula = 6  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoChartFieldPercentage = 3  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoChartFieldRange = 7  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoChartFieldSeriesName = 4  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoChartFieldValue = 5  # Enum: MsoChartFieldType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msochartfieldtype
msoClipboardFormatHTML = 2  # Enum: MsoClipboardFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoclipboardformat
msoClipboardFormatMixed = (
    -2
)  # Enum: MsoClipboardFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoclipboardformat
msoClipboardFormatNative = 1  # Enum: MsoClipboardFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoclipboardformat
msoClipboardFormatPlainText = 4  # Enum: MsoClipboardFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoclipboardformat
msoClipboardFormatRTF = 3  # Enum: MsoClipboardFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoclipboardformat
msoColorTypeCMS = 4  # Enum: MsoColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocolortype
msoColorTypeCMYK = 3  # Enum: MsoColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocolortype
msoColorTypeInk = 5  # Enum: MsoColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocolortype
msoColorTypeMixed = (
    -2
)  # Enum: MsoColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocolortype
msoColorTypeRGB = 1  # Enum: MsoColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocolortype
msoColorTypeScheme = 2  # Enum: MsoColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocolortype
msoComboLabel = 1  # Enum: MsoComboStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocombostyle
msoComboNormal = 0  # Enum: MsoComboStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocombostyle
msoCommandBarButtonHyperlinkInsertPicture = 2  # Enum: MsoCommandBarButtonHyperlinkType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocommandbarbuttonhyperlinktype
msoCommandBarButtonHyperlinkNone = 0  # Enum: MsoCommandBarButtonHyperlinkType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocommandbarbuttonhyperlinktype
msoCommandBarButtonHyperlinkOpen = 1  # Enum: MsoCommandBarButtonHyperlinkType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocommandbarbuttonhyperlinktype
msoConditionAnyNumberBetween = 34  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionAnytime = 25  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionAnytimeBetween = 26  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionAtLeast = 36  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionAtMost = 35  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionBeginsWith = 11  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionDoesNotEqual = 33  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEndsWith = 12  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEquals = 32  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsCompleted = 66  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsDeferred = 68  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsHigh = 60  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsInProgress = 65  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsLow = 58  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsNormal = 59  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsNotStarted = 64  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionEqualsWaitingForSomeoneElse = 67  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeAllFiles = 1  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeBinders = 6  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeCalendarItem = 45  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeContactItem = 46  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeDataConnectionFiles = 51  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeDatabases = 7  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeDesignerFiles = 56  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeDocumentImagingFiles = 54  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeExcelWorkbooks = 4  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeJournalItem = 48  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeMailItem = 44  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeNoteItem = 47  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeOfficeFiles = 2  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeOutlookItems = 43  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypePhotoDrawFiles = 50  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypePowerPointPresentations = 5  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeProjectFiles = 53  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypePublisherFiles = 52  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeTaskItem = 49  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeTemplates = 8  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeVisioFiles = 55  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeWebPages = 57  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFileTypeWordDocuments = 3  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionFreeText = 42  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionInTheLast = 31  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionInTheNext = 30  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIncludes = 9  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIncludesFormsOf = 41  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIncludesNearEachOther = 13  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIncludesPhrase = 10  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIsExactly = 14  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIsNo = 40  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIsNot = 15  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionIsYes = 39  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionLastMonth = 22  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionLastWeek = 19  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionLessThan = 38  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionMoreThan = 37  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNextMonth = 24  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNextWeek = 21  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToCompleted = 71  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToDeferred = 73  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToHigh = 63  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToInProgress = 70  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToLow = 61  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToNormal = 62  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToNotStarted = 69  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionNotEqualToWaitingForSomeoneElse = 72  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionOn = 27  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionOnOrAfter = 28  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionOnOrBefore = 29  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionThisMonth = 23  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionThisWeek = 20  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionToday = 17  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionTomorrow = 18  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConditionYesterday = 16  # Enum: MsoCondition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocondition
msoConnectorAnd = 1  # Enum: MsoConnector, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoconnector
msoConnectorOr = 2  # Enum: MsoConnector, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoconnector
msoConnectorCurve = 3  # Enum: MsoConnectorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
msoConnectorElbow = 2  # Enum: MsoConnectorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
msoConnectorStraight = 1  # Enum: MsoConnectorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
msoConnectorTypeMixed = (
    -2
)  # Enum: MsoConnectorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoconnectortype
msoContactCardAddressTypeIM = 3  # Enum: MsoContactCardAddressType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardaddresstype
msoContactCardAddressTypeOutlook = 1  # Enum: MsoContactCardAddressType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardaddresstype
msoContactCardAddressTypeSMTP = 2  # Enum: MsoContactCardAddressType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardaddresstype
msoContactCardAddressTypeUnknown = 0  # Enum: MsoContactCardAddressType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardaddresstype
msoContactCardFull = 1  # Enum: MsoContactCardStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardstyle
msoContactCardHover = 0  # Enum: MsoContactCardStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardstyle
msoContactCardTypeEnterpriseContact = 0  # Enum: MsoContactCardType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardtype
msoContactCardTypeEnterpriseGroup = 3  # Enum: MsoContactCardType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardtype
msoContactCardTypePersonalContact = 1  # Enum: MsoContactCardType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardtype
msoContactCardTypePersonalDistributionList = 4  # Enum: MsoContactCardType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardtype
msoContactCardTypeUnknownContact = 2  # Enum: MsoContactCardType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontactcardtype
msoControlOLEUsageBoth = 3  # Enum: MsoControlOLEUsage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontrololeusage
msoControlOLEUsageClient = 2  # Enum: MsoControlOLEUsage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontrololeusage
msoControlOLEUsageNeither = 0  # Enum: MsoControlOLEUsage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontrololeusage
msoControlOLEUsageServer = 1  # Enum: MsoControlOLEUsage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontrololeusage
msoControlActiveX = 22  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlAutoCompleteCombo = 26  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlButton = 1  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlButtonDropdown = 5  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlButtonPopup = 12  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlComboBox = 4  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlCustom = 0  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlDropdown = 3  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlEdit = 2  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlExpandingGrid = 16  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlGauge = 19  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlGenericDropdown = 8  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlGraphicCombo = 20  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlGraphicDropdown = 9  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlGraphicPopup = 11  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlGrid = 18  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlLabel = 15  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlLabelEx = 24  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlOCXDropdown = 7  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlPane = 21  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlPopup = 10  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlSpinner = 23  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlSplitButtonMRUPopup = 14  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlSplitButtonPopup = 13  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlSplitDropdown = 6  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlSplitExpandingGrid = 17  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoControlWorkPane = 25  # Enum: MsoControlType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocontroltype
msoCustomXMLNodeAttribute = 2  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLNodeCData = 4  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLNodeComment = 8  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLNodeDocument = 9  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLNodeElement = 1  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLNodeProcessingInstruction = 7  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLNodeText = 3  # Enum: MsoCustomXMLNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlnodetype
msoCustomXMLValidationErrorAutomaticallyCleared = 1  # Enum: MsoCustomXMLValidationErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlvalidationerrortype
msoCustomXMLValidationErrorManual = 2  # Enum: MsoCustomXMLValidationErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlvalidationerrortype
msoCustomXMLValidationErrorSchemaGenerated = 0  # Enum: MsoCustomXMLValidationErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msocustomxmlvalidationerrortype
msoDateTimeFigureOut = 14  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeFormatMixed = (
    -2
)  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeHmm = 10  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeHmmss = 11  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeMMMMdyyyy = 4  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeMMMMyy = 6  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeMMddyyHmm = 8  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeMMddyyhmmAMPM = 9  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeMMyy = 7  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeMdyy = 1  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimedMMMMyyyy = 3  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimedMMMyy = 5  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimeddddMMMMddyyyy = 2  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimehmmAMPM = 12  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDateTimehmmssAMPM = 13  # Enum: MsoDateTimeFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodatetimeformat
msoDiagramAssistant = 2  # Enum: MsoDiagramNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramnodetype
msoDiagramNode = 1  # Enum: MsoDiagramNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramnodetype
msoDiagramCycle = 2  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDiagramMixed = (
    -2
)  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDiagramOrgChart = 1  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDiagramPyramid = 4  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDiagramRadial = 3  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDiagramTarget = 6  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDiagramVenn = 5  # Enum: MsoDiagramType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodiagramtype
msoDistributeHorizontally = 0  # Enum: MsoDistributeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodistributecmd
msoDistributeVertically = 1  # Enum: MsoDistributeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodistributecmd
msoDocInspectorStatusDocOk = 0  # Enum: MsoDocInspectorStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocinspectorstatus
msoDocInspectorStatusError = 2  # Enum: MsoDocInspectorStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocinspectorstatus
msoDocInspectorStatusIssueFound = 1  # Enum: MsoDocInspectorStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocinspectorstatus
msoPropertyTypeBoolean = 2  # Enum: MsoDocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocproperties
msoPropertyTypeDate = 3  # Enum: MsoDocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocproperties
msoPropertyTypeFloat = 5  # Enum: MsoDocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocproperties
msoPropertyTypeNumber = 1  # Enum: MsoDocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocproperties
msoPropertyTypeString = 4  # Enum: MsoDocProperties, https://docs.microsoft.com/ja-jp/office/vba/api/office.msodocproperties
msoEditingAuto = 0  # Enum: MsoEditingType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoeditingtype
msoEditingCorner = 1  # Enum: MsoEditingType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoeditingtype
msoEditingSmooth = 2  # Enum: MsoEditingType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoeditingtype
msoEditingSymmetric = 3  # Enum: MsoEditingType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoeditingtype
msoEncodingArabic = 1256  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingArabicASMO = 708  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingArabicAutoDetect = 51256  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingArabicTransparentASMO = 720  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingAutoDetect = 50001  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingBaltic = 1257  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingCentralEuropean = 1250  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingCyrillic = 1251  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingCyrillicAutoDetect = 51251  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICArabic = 20420  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICDenmarkNorway = 20277  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICFinlandSweden = 20278  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICFrance = 20297  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICGermany = 20273  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICGreek = 20423  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICGreekModern = 875  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICHebrew = 20424  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICIcelandic = 20871  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICInternational = 500  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICItaly = 20280  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICJapaneseKatakanaExtended = 20290  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICKoreanExtended = 20833  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICKoreanExtendedAndKorean = 50933  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICLatinAmericaSpain = 20284  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICMultilingualROECELatin2 = 870  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICRussian = 20880  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICSerbianBulgarian = 21025  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICThai = 20838  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICTurkish = 20905  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICTurkishLatin5 = 1026  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICUSCanada = 37  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICUSCanadaAndJapanese = 50931  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICUSCanadaAndTraditionalChinese = 50937  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEBCDICUnitedKingdom = 20285  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEUCChineseSimplifiedChinese = 51936  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEUCJapanese = 51932  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEUCKorean = 51949  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEUCTaiwaneseTraditionalChinese = 51950  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingEuropa3 = 29001  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingExtAlphaLowercase = 21027  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingGreek = 1253  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingGreekAutoDetect = 51253  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingHZGBSimplifiedChinese = 52936  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingHebrew = 1255  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingIA5German = 20106  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingIA5IRV = 20105  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingIA5Norwegian = 20108  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingIA5Swedish = 20107  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIAssamese = 57006  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIBengali = 57003  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIDevanagari = 57002  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIGujarati = 57010  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIKannada = 57008  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIMalayalam = 57009  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIOriya = 57007  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIIPunjabi = 57011  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIITamil = 57004  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISCIITelugu = 57005  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO2022CNSimplifiedChinese = 50229  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO2022CNTraditionalChinese = 50227  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO2022JPJISX02011989 = 50222  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO2022JPJISX02021984 = 50221  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO2022JPNoHalfwidthKatakana = 50220  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO2022KR = 50225  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO6937NonSpacingAccent = 20269  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO885915Latin9 = 28605  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88591Latin1 = 28591  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88592CentralEurope = 28592  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88593Latin3 = 28593  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88594Baltic = 28594  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88595Cyrillic = 28595  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88596Arabic = 28596  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88597Greek = 28597  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88598Hebrew = 28598  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88598HebrewLogical = 38598  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingISO88599Turkish = 28599  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingJapaneseAutoDetect = 50932  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingJapaneseShiftJIS = 932  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingKOI8R = 20866  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingKOI8U = 21866  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingKorean = 949  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingKoreanAutoDetect = 50949  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingKoreanJohab = 1361  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacArabic = 10004  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacCroatia = 10082  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacCyrillic = 10007  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacGreek1 = 10006  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacHebrew = 10005  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacIcelandic = 10079  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacJapanese = 10001  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacKorean = 10003  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacLatin2 = 10029  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacRoman = 10000  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacRomania = 10010  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacSimplifiedChineseGB2312 = 10008  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacTraditionalChineseBig5 = 10002  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacTurkish = 10081  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingMacUkraine = 10017  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMArabic = 864  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMBaltic = 775  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMCanadianFrench = 863  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMCyrillic = 855  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMCyrillicII = 866  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMGreek437G = 737  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMHebrew = 862  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMIcelandic = 861  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMModernGreek = 869  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMMultilingualLatinI = 850  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMMultilingualLatinII = 852  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMNordic = 865  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMPortuguese = 860  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMTurkish = 857  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingOEMUnitedStates = 437  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingSimplifiedChineseAutoDetect = 50936  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingSimplifiedChineseGB18030 = 54936  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingSimplifiedChineseGBK = 936  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingT61 = 20261  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTaiwanCNS = 20000  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTaiwanEten = 20002  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTaiwanIBM5550 = 20003  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTaiwanTCA = 20001  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTaiwanTeleText = 20004  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTaiwanWang = 20005  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingThai = 874  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTraditionalChineseAutoDetect = 50950  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTraditionalChineseBig5 = 950  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingTurkish = 1254  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingUSASCII = 20127  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingUTF7 = 65000  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingUTF8 = 65001  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingUnicodeBigEndian = 1201  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingUnicodeLittleEndian = 1200  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingVietnamese = 1258  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoEncodingWestern = 1252  # Enum: MsoEncoding, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoencoding
msoMethodGet = 0  # Enum: MsoExtraInfoMethod, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoextrainfomethod
msoMethodPost = 1  # Enum: MsoExtraInfoMethod, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoextrainfomethod
msoExtrusionColorAutomatic = 1  # Enum: MsoExtrusionColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoextrusioncolortype
msoExtrusionColorCustom = 2  # Enum: MsoExtrusionColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoextrusioncolortype
msoExtrusionColorTypeMixed = (
    -2
)  # Enum: MsoExtrusionColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoextrusioncolortype
MsoFarEastLineBreakLanguageJapanese = 1041  # Enum: MsoFarEastLineBreakLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofareastlinebreaklanguageid
MsoFarEastLineBreakLanguageKorean = 1042  # Enum: MsoFarEastLineBreakLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofareastlinebreaklanguageid
MsoFarEastLineBreakLanguageSimplifiedChinese = 2052  # Enum: MsoFarEastLineBreakLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofareastlinebreaklanguageid
MsoFarEastLineBreakLanguageTraditionalChinese = 1028  # Enum: MsoFarEastLineBreakLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofareastlinebreaklanguageid
msoFeatureInstallNone = 0  # Enum: MsoFeatureInstall, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofeatureinstall
msoFeatureInstallOnDemand = 1  # Enum: MsoFeatureInstall, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofeatureinstall
msoFeatureInstallOnDemandWithUI = 2  # Enum: MsoFeatureInstall, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofeatureinstall
msoFileDialogFilePicker = 3  # Enum: MsoFileDialogType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogtype
msoFileDialogFolderPicker = 4  # Enum: MsoFileDialogType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogtype
msoFileDialogOpen = 1  # Enum: MsoFileDialogType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogtype
msoFileDialogSaveAs = 2  # Enum: MsoFileDialogType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogtype
msoFileDialogViewDetails = 2  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewLargeIcons = 6  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewList = 1  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewPreview = 4  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewProperties = 3  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewSmallIcons = 7  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewThumbnail = 5  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewTiles = 9  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoFileDialogViewWebView = 8  # Enum: MsoFileDialogView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiledialogview
msoListbyName = 1  # Enum: MsoFileFindListBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindlistby
msoListbyTitle = 2  # Enum: MsoFileFindListBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindlistby
msoOptionsAdd = 2  # Enum: MsoFileFindOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindoptions
msoOptionsNew = 1  # Enum: MsoFileFindOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindoptions
msoOptionsWithin = 3  # Enum: MsoFileFindOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindoptions
msoFileFindSortbyAuthor = 1  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoFileFindSortbyDateCreated = 2  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoFileFindSortbyDateSaved = 4  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoFileFindSortbyFileName = 5  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoFileFindSortbyLastSavedBy = 3  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoFileFindSortbySize = 6  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoFileFindSortbyTitle = 7  # Enum: MsoFileFindSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindsortby
msoViewFileInfo = 1  # Enum: MsoFileFindView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindview
msoViewPreview = 2  # Enum: MsoFileFindView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindview
msoViewSummaryInfo = 3  # Enum: MsoFileFindView, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilefindview
msoCreateNewFile = 1  # Enum: MsoFileNewAction, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewaction
msoEditFile = 0  # Enum: MsoFileNewAction, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewaction
msoOpenFile = 2  # Enum: MsoFileNewAction, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewaction
msoBottomSection = 4  # Enum: MsoFileNewSection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewsection
msoNew = 1  # Enum: MsoFileNewSection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewsection
msoNewfromExistingFile = 2  # Enum: MsoFileNewSection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewsection
msoNewfromTemplate = 3  # Enum: MsoFileNewSection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewsection
msoOpenDocument = 0  # Enum: MsoFileNewSection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilenewsection
msoFileTypeAllFiles = 1  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeBinders = 6  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeCalendarItem = 11  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeContactItem = 12  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeDataConnectionFiles = 17  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeDatabases = 7  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeDesignerFiles = 22  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeDocumentImagingFiles = 20  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeExcelWorkbooks = 4  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeJournalItem = 14  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeMailItem = 10  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeNoteItem = 13  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeOfficeFiles = 2  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeOutlookItems = 9  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypePhotoDrawFiles = 16  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypePowerPointPresentations = 5  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeProjectFiles = 19  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypePublisherFiles = 18  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeTaskItem = 15  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeTemplates = 8  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeVisioFiles = 21  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeWebPages = 23  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileTypeWordDocuments = 3  # Enum: MsoFileType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiletype
msoFileValidationDefault = 0  # Enum: MsoFileValidationMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilevalidationmode
msoFileValidationSkip = 1  # Enum: MsoFileValidationMode, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilevalidationmode
msoFillBackground = 5  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFillGradient = 3  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFillMixed = (
    -2
)  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFillPatterned = 2  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFillPicture = 6  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFillSolid = 1  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFillTextured = 4  # Enum: MsoFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilltype
msoFilterComparisonContains = 8  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonEqual = 0  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonGreaterThan = 3  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonGreaterThanEqual = 5  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonIsBlank = 6  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonIsNotBlank = 7  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonLessThan = 2  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonLessThanEqual = 4  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonNotContains = 9  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterComparisonNotEqual = 1  # Enum: MsoFilterComparison, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofiltercomparison
msoFilterConjunctionAnd = 0  # Enum: MsoFilterConjunction, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilterconjunction
msoFilterConjunctionOr = 1  # Enum: MsoFilterConjunction, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofilterconjunction
msoFlipHorizontal = 0  # Enum: MsoFlipCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoflipcmd
msoFlipVertical = 1  # Enum: MsoFlipCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoflipcmd
msoThemeComplexScript = 2  # Enum: MsoFontLanguageIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofontlanguageindex
msoThemeEastAsian = 3  # Enum: MsoFontLanguageIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofontlanguageindex
msoThemeLatin = 1  # Enum: MsoFontLanguageIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msofontlanguageindex
msoGradientColorMixed = (
    -2
)  # Enum: MsoGradientColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientcolortype
msoGradientMultiColor = 4  # Enum: MsoGradientColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientcolortype
msoGradientOneColor = 1  # Enum: MsoGradientColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientcolortype
msoGradientPresetColors = 3  # Enum: MsoGradientColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientcolortype
msoGradientTwoColors = 2  # Enum: MsoGradientColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientcolortype
msoGradientDiagonalDown = 4  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientDiagonalUp = 3  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientFromCenter = 7  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientFromCorner = 5  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientFromTitle = 6  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientHorizontal = 1  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientMixed = (
    -2
)  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGradientVertical = 2  # Enum: MsoGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msogradientstyle
msoGraphicStyleMixed = (
    -2
)  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStyleNotAPreset = 0  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset1 = 1  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset10 = 10  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset11 = 11  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset12 = 12  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset13 = 13  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset14 = 14  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset15 = 15  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset16 = 16  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset17 = 17  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset18 = 18  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset19 = 19  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset2 = 2  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset20 = 20  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset21 = 21  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset22 = 22  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset23 = 23  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset24 = 24  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset25 = 25  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset26 = 26  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset27 = 27  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset28 = 28  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset3 = 3  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset4 = 4  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset5 = 5  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset6 = 6  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset7 = 7  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset8 = 8  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoGraphicStylePreset9 = 9  # Enum: MsoGraphicStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msographicstyleindex
msoHTMLProjectOpenSourceView = 1  # Enum: MsoHTMLProjectOpen, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohtmlprojectopen
msoHTMLProjectOpenTextView = 2  # Enum: MsoHTMLProjectOpen, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohtmlprojectopen
msoHTMLProjectStateDocumentLocked = 1  # Enum: MsoHTMLProjectState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohtmlprojectstate
msoHTMLProjectStateDocumentProjectUnlocked = 3  # Enum: MsoHTMLProjectState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohtmlprojectstate
msoHTMLProjectStateProjectLocked = 2  # Enum: MsoHTMLProjectState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohtmlprojectstate
msoAnchorCenter = 2  # Enum: MsoHorizontalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohorizontalanchor
msoAnchorNone = 1  # Enum: MsoHorizontalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohorizontalanchor
msoHorizontalAnchorMixed = (
    -2
)  # Enum: MsoHorizontalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohorizontalanchor
msoHyperlinkInlineShape = 2  # Enum: MsoHyperlinkType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohyperlinktype
msoHyperlinkRange = 0  # Enum: MsoHyperlinkType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohyperlinktype
msoHyperlinkShape = 1  # Enum: MsoHyperlinkType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msohyperlinktype
msoIconAlert = 2  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIconAlertCritical = 7  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIconAlertInfo = 4  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIconAlertQuery = 6  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIconAlertWarning = 5  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIconNone = 0  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIconTip = 3  # Enum: MsoIconType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoicontype
msoIodGroupPIAs = 0  # Enum: MsoIodGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoiodgroup
msoIodGroupVSTOR35Mgd = 1  # Enum: MsoIodGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoiodgroup
msoIodGroupVSTOR40Mgd = 2  # Enum: MsoIodGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoiodgroup
msoLanguageIDAfrikaans = 1078  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDAlbanian = 1052  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDAmharic = 1118  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabic = 1025  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicAlgeria = 5121  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicBahrain = 15361  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicEgypt = 3073  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicIraq = 2049  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicJordan = 11265  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicKuwait = 13313  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicLebanon = 12289  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicLibya = 4097  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicMorocco = 6145  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicOman = 8193  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicQatar = 16385  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicSyria = 10241  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicTunisia = 7169  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicUAE = 14337  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArabicYemen = 9217  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDArmenian = 1067  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDAssamese = 1101  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDAzeriCyrillic = 2092  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDAzeriLatin = 1068  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBasque = 1069  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBelgianDutch = 2067  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBelgianFrench = 2060  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBengali = 1093  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBosnian = 4122  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBosnianBosniaHerzegovinaCyrillic = 8218  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBosnianBosniaHerzegovinaLatin = 5146  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBrazilianPortuguese = 1046  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBulgarian = 1026  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDBurmese = 1109  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDByelorussian = 1059  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDCatalan = 1027  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDCherokee = 1116  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDChineseHongKongSAR = 3076  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDChineseMacaoSAR = 5124  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDChineseSingapore = 4100  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDCroatian = 1050  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDCzech = 1029  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDDanish = 1030  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDDivehi = 1125  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDDutch = 1043  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDDzongkhaBhutan = 2129  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEdo = 1126  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishAUS = 3081  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishBelize = 10249  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishCanadian = 4105  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishCaribbean = 9225  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishIndonesia = 14345  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishIreland = 6153  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishJamaica = 8201  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishNewZealand = 5129  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishPhilippines = 13321  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishSouthAfrica = 7177  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishTrinidadTobago = 11273  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishUK = 2057  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishUS = 1033  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEnglishZimbabwe = 12297  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDEstonian = 1061  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFaeroese = 1080  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFarsi = 1065  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFilipino = 1124  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFinnish = 1035  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrench = 1036  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchCameroon = 11276  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchCanadian = 3084  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchCongoDRC = 9228  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchCotedIvoire = 12300  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchHaiti = 15372  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchLuxembourg = 5132  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchMali = 13324  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchMonaco = 6156  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchMorocco = 14348  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchReunion = 8204  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchSenegal = 10252  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchWestIndies = 7180  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrenchZaire = 9228  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFrisianNetherlands = 1122  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDFulfulde = 1127  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGaelicIreland = 2108  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGaelicScotland = 1084  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGalician = 1110  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGeorgian = 1079  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGerman = 1031  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGermanAustria = 3079  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGermanLiechtenstein = 5127  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGermanLuxembourg = 4103  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGreek = 1032  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGuarani = 1140  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDGujarati = 1095  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDHausa = 1128  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDHawaiian = 1141  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDHebrew = 1037  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDHindi = 1081  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDHungarian = 1038  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDIbibio = 1129  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDIcelandic = 1039  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDIgbo = 1136  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDIndonesian = 1057  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDInuktitut = 1117  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDItalian = 1040  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDJapanese = 1041  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKannada = 1099  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKanuri = 1137  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKashmiri = 1120  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKashmiriDevanagari = 2144  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKazakh = 1087  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKhmer = 1107  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKirghiz = 1088  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKonkani = 1111  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKorean = 1042  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDKyrgyz = 1088  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDLao = 1108  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDLatin = 1142  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDLatvian = 1062  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDLithuanian = 1063  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMacedonian = 1071  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMacedonianFYROM = 1071  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMalayBruneiDarussalam = 2110  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMalayalam = 1100  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMalaysian = 1086  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMaltese = 1082  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDManipuri = 1112  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMaori = 1153  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMarathi = 1102  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMexicanSpanish = 2058  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMixed = (
    -2
)  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDMongolian = 1104  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDNepali = 1121  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDNoProofing = 1024  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDNone = 0  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDNorwegianBokmol = 1044  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDNorwegianNynorsk = 2068  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDOriya = 1096  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDOromo = 1138  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDPashto = 1123  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDPolish = 1045  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDPortuguese = 2070  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDPunjabi = 1094  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDQuechuaBolivia = 1131  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDQuechuaEcuador = 2155  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDQuechuaPeru = 3179  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDRhaetoRomanic = 1047  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDRomanian = 1048  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDRomanianMoldova = 2072  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDRussian = 1049  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDRussianMoldova = 2073  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSamiLappish = 1083  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSanskrit = 1103  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSepedi = 1132  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSerbianBosniaHerzegovinaCyrillic = 7194  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSerbianBosniaHerzegovinaLatin = 6170  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSerbianCyrillic = 3098  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSerbianLatin = 2074  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSesotho = 1072  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSimplifiedChinese = 2052  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSindhi = 1113  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSindhiPakistan = 2137  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSinhalese = 1115  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSlovak = 1051  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSlovenian = 1060  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSomali = 1143  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSorbian = 1070  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanish = 1034  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishArgentina = 11274  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishBolivia = 16394  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishChile = 13322  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishColombia = 9226  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishCostaRica = 5130  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishDominicanRepublic = 7178  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishEcuador = 12298  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishElSalvador = 17418  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishGuatemala = 4106  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishHonduras = 18442  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishModernSort = 3082  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishNicaragua = 19466  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishPanama = 6154  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishParaguay = 15370  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishPeru = 10250  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishPuertoRico = 20490  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishUruguay = 14346  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSpanishVenezuela = 8202  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSutu = 1072  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSwahili = 1089  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSwedish = 1053  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSwedishFinland = 2077  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSwissFrench = 4108  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSwissGerman = 2055  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSwissItalian = 2064  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDSyriac = 1114  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTajik = 1064  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTamazight = 1119  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTamazightLatin = 2143  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTamil = 1097  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTatar = 1092  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTelugu = 1098  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDThai = 1054  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTibetan = 1105  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTigrignaEritrea = 2163  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTigrignaEthiopic = 1139  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTraditionalChinese = 1028  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTsonga = 1073  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTswana = 1074  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTurkish = 1055  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDTurkmen = 1090  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDUkrainian = 1058  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDUrdu = 1056  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDUzbekCyrillic = 2115  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDUzbekLatin = 1091  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDVenda = 1075  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDVietnamese = 1066  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDWelsh = 1106  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDXhosa = 1076  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDYi = 1144  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDYiddish = 1085  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDYoruba = 1130  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDZulu = 1077  # Enum: MsoLanguageID, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageid
msoLanguageIDChineseHongKong = 3076  # Enum: MsoLanguageIDHidden, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageidhidden
msoLanguageIDChineseMacao = 5124  # Enum: MsoLanguageIDHidden, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageidhidden
msoLanguageIDEnglishTrinidad = 11273  # Enum: MsoLanguageIDHidden, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolanguageidhidden
msoLastModifiedAnyTime = 7  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLastModifiedLastMonth = 5  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLastModifiedLastWeek = 3  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLastModifiedThisMonth = 6  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLastModifiedThisWeek = 4  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLastModifiedToday = 2  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLastModifiedYesterday = 1  # Enum: MsoLastModified, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolastmodified
msoLightRigBalanced = 14  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigBrightRoom = 27  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigChilly = 22  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigContrasting = 18  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigFlat = 24  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigFlood = 17  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigFreezing = 23  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigGlow = 26  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigHarsh = 16  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyFlat1 = 1  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyFlat2 = 2  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyFlat3 = 3  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyFlat4 = 4  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyHarsh1 = 9  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyHarsh2 = 10  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyHarsh3 = 11  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyHarsh4 = 12  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyNormal1 = 5  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyNormal2 = 6  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyNormal3 = 7  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigLegacyNormal4 = 8  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigMixed = (
    -2
)  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigMorning = 19  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigSoft = 15  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigSunrise = 20  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigSunset = 21  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigThreePoint = 13  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLightRigTwoPoint = 25  # Enum: MsoLightRigType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolightrigtype
msoLineCapFlat = 3  # Enum: MsoLineCapStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinecapstyle
msoLineCapMixed = (
    -2
)  # Enum: MsoLineCapStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinecapstyle
msoLineCapRound = 2  # Enum: MsoLineCapStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinecapstyle
msoLineCapSquare = 1  # Enum: MsoLineCapStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinecapstyle
msoLineDash = 4  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineDashDot = 5  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineDashDotDot = 6  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineDashStyleMixed = (
    -2
)  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineLongDash = 7  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineLongDashDot = 8  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineLongDashDotDot = 9  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineRoundDot = 3  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineSolid = 1  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineSquareDot = 2  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineSysDash = 10  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineSysDashDot = 12  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineSysDot = 11  # Enum: MsoLineDashStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinedashstyle
msoLineFillBackground = 5  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillGradient = 3  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillMixed = (
    -2
)  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillNone = 0  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillPatterned = 2  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillPicture = 6  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillSolid = 1  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineFillTextured = 4  # Enum: MsoLineFillType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinefilltype
msoLineJoinBevel = 2  # Enum: MsoLineJoinStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinejoinstyle
msoLineJoinMiter = 3  # Enum: MsoLineJoinStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinejoinstyle
msoLineJoinMixed = (
    -2
)  # Enum: MsoLineJoinStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinejoinstyle
msoLineJoinRound = 1  # Enum: MsoLineJoinStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinejoinstyle
msoLineSingle = 1  # Enum: MsoLineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinestyle
msoLineStyleMixed = (
    -2
)  # Enum: MsoLineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinestyle
msoLineThickBetweenThin = 5  # Enum: MsoLineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinestyle
msoLineThickThin = 4  # Enum: MsoLineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinestyle
msoLineThinThick = 3  # Enum: MsoLineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinestyle
msoLineThinThin = 2  # Enum: MsoLineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msolinestyle
msoMenuAnimationNone = 0  # Enum: MsoMenuAnimation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomenuanimation
msoMenuAnimationRandom = 1  # Enum: MsoMenuAnimation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomenuanimation
msoMenuAnimationSlide = 3  # Enum: MsoMenuAnimation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomenuanimation
msoMenuAnimationUnfold = 2  # Enum: MsoMenuAnimation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomenuanimation
msoMergeCombine = 2  # Enum: MsoMergeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomergecmd
msoMergeFragment = 5  # Enum: MsoMergeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomergecmd
msoMergeIntersect = 3  # Enum: MsoMergeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomergecmd
msoMergeSubtract = 4  # Enum: MsoMergeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomergecmd
msoMergeUnion = 1  # Enum: MsoMergeCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomergecmd
msoMetaPropertyTypeBoolean = 1  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeBusinessData = 20  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeBusinessDataSecondary = 21  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeCalculated = 3  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeChoice = 2  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeComputed = 4  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeCurrency = 5  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeDateTime = 6  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeFillInChoice = 7  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeGuid = 8  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeInteger = 9  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeLookup = 10  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeMax = 22  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeMultiChoice = 12  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeMultiChoiceFillIn = 13  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeMultiChoiceLookup = 11  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeNote = 14  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeNumber = 15  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeText = 16  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeUnknown = 0  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeUrl = 17  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeUser = 18  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoMetaPropertyTypeUserMulti = 19  # Enum: MsoMetaPropertyType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msometapropertytype
msoIntegerMixed = 32768  # Enum: MsoMixedType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomixedtype
msoSingleMixed = (
    -2147483648
)  # Enum: MsoMixedType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomixedtype
msoModeAutoDown = 1  # Enum: MsoModeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomodetype
msoModeModal = 0  # Enum: MsoModeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomodetype
msoModeModeless = 2  # Enum: MsoModeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomodetype
msoMoveRowFirst = (
    -4
)  # Enum: MsoMoveRow, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomoverow
msoMoveRowNbr = (
    -1
)  # Enum: MsoMoveRow, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomoverow
msoMoveRowNext = (
    -2
)  # Enum: MsoMoveRow, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomoverow
msoMoveRowPrev = (
    -3
)  # Enum: MsoMoveRow, https://docs.microsoft.com/ja-jp/office/vba/api/office.msomoverow
msoBulletAlphaLCParenBoth = 8  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletAlphaLCParenRight = 9  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletAlphaLCPeriod = 0  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletAlphaUCParenBoth = 10  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletAlphaUCParenRight = 11  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletAlphaUCPeriod = 1  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicAbjadDash = 24  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicAlphaDash = 23  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicDBPeriod = 29  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicDBPlain = 28  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicParenBoth = 12  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicParenRight = 2  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicPeriod = 3  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletArabicPlain = 13  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletCircleNumDBPlain = 18  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletCircleNumWDBlackPlain = 20  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletCircleNumWDWhitePlain = 19  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletHebrewAlphaDash = 25  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletHindiAlpha1Period = 40  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletHindiAlphaPeriod = 36  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletHindiNumParenRight = 39  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletHindiNumPeriod = 37  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletKanjiKoreanPeriod = 27  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletKanjiKoreanPlain = 26  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletKanjiSimpChinDBPeriod = 38  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletRomanLCParenBoth = 4  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletRomanLCParenRight = 5  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletRomanLCPeriod = 6  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletRomanUCParenBoth = 14  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletRomanUCParenRight = 15  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletRomanUCPeriod = 7  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletSimpChinPeriod = 17  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletSimpChinPlain = 16  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletStyleMixed = (
    -2
)  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletThaiAlphaParenBoth = 32  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletThaiAlphaParenRight = 31  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletThaiAlphaPeriod = 30  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletThaiNumParenBoth = 35  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletThaiNumParenRight = 34  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletThaiNumPeriod = 33  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletTradChinPeriod = 22  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoBulletTradChinPlain = 21  # Enum: MsoNumberedBulletStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msonumberedbulletstyle
msoOLEMenuGroupContainer = 2  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOLEMenuGroupEdit = 1  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOLEMenuGroupFile = 0  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOLEMenuGroupHelp = 5  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOLEMenuGroupNone = (
    -1
)  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOLEMenuGroupObject = 3  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOLEMenuGroupWindow = 4  # Enum: MsoOLEMenuGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoolemenugroup
msoOrgChartLayoutBothHanging = 2  # Enum: MsoOrgChartLayoutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartlayouttype
msoOrgChartLayoutDefault = 5  # Enum: MsoOrgChartLayoutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartlayouttype
msoOrgChartLayoutLeftHanging = 3  # Enum: MsoOrgChartLayoutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartlayouttype
msoOrgChartLayoutMixed = (
    -2
)  # Enum: MsoOrgChartLayoutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartlayouttype
msoOrgChartLayoutRightHanging = 4  # Enum: MsoOrgChartLayoutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartlayouttype
msoOrgChartLayoutStandard = 1  # Enum: MsoOrgChartLayoutType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartlayouttype
msoOrgChartOrientationMixed = (
    -2
)  # Enum: MsoOrgChartOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartorientation
msoOrgChartOrientationVertical = 1  # Enum: MsoOrgChartOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorgchartorientation
msoOrientationHorizontal = 1  # Enum: MsoOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorientation
msoOrientationMixed = (
    -2
)  # Enum: MsoOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorientation
msoOrientationVertical = 2  # Enum: MsoOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoorientation
msoAlignCenter = 2  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignDistribute = 5  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignJustify = 4  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignJustifyLow = 7  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignLeft = 1  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignMixed = (
    -2
)  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignRight = 3  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoAlignThaiDistribute = 6  # Enum: MsoParagraphAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoparagraphalignment
msoPathType1 = 1  # Enum: MsoPathFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopathformat
msoPathType2 = 2  # Enum: MsoPathFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopathformat
msoPathType3 = 3  # Enum: MsoPathFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopathformat
msoPathType4 = 4  # Enum: MsoPathFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopathformat
msoPathTypeMixed = (
    -2
)  # Enum: MsoPathFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopathformat
msoPathTypeNone = 0  # Enum: MsoPathFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopathformat
msoPattern10Percent = 2  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern20Percent = 3  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern25Percent = 4  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern30Percent = 5  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern40Percent = 6  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern50Percent = 7  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern5Percent = 1  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern60Percent = 8  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern70Percent = 9  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern75Percent = 10  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern80Percent = 11  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPattern90Percent = 12  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternCross = 51  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDarkDownwardDiagonal = 15  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDarkHorizontal = 13  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDarkUpwardDiagonal = 16  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDarkVertical = 14  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDashedDownwardDiagonal = 28  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDashedHorizontal = 32  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDashedUpwardDiagonal = 27  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDashedVertical = 31  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDiagonalBrick = 40  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDiagonalCross = 54  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDivot = 46  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDottedDiamond = 24  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDottedGrid = 45  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternDownwardDiagonal = 52  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternHorizontal = 49  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternHorizontalBrick = 35  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLargeCheckerBoard = 36  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLargeConfetti = 33  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLargeGrid = 34  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLightDownwardDiagonal = 21  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLightHorizontal = 19  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLightUpwardDiagonal = 22  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternLightVertical = 20  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternMixed = (
    -2
)  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternNarrowHorizontal = 30  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternNarrowVertical = 29  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternOutlinedDiamond = 41  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternPlaid = 42  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternShingle = 47  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternSmallCheckerBoard = 17  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternSmallConfetti = 37  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternSmallGrid = 23  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternSolidDiamond = 39  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternSphere = 43  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternTrellis = 18  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternUpwardDiagonal = 53  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternVertical = 50  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternWave = 48  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternWeave = 44  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternWideDownwardDiagonal = 25  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternWideUpwardDiagonal = 26  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPatternZigZag = 38  # Enum: MsoPatternType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopatterntype
msoPermissionAllCommon = 127  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionChange = 15  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionEdit = 2  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionExtract = 8  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionFullControl = 64  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionObjModel = 32  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionPrint = 16  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionRead = 1  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionSave = 4  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPermissionView = 1  # Enum: MsoPermission, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopermission
msoPickerFieldDateTime = 1  # Enum: MsoPickerField, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopickerfield
msoPickerFieldMax = 5  # Enum: MsoPickerField, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopickerfield
msoPickerFieldNumber = 2  # Enum: MsoPickerField, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopickerfield
msoPickerFieldText = 3  # Enum: MsoPickerField, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopickerfield
msoPickerFieldUnknown = 0  # Enum: MsoPickerField, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopickerfield
msoPickerFieldUser = 4  # Enum: MsoPickerField, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopickerfield
msoPictureAutomatic = 1  # Enum: MsoPictureColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecolortype
msoPictureBlackAndWhite = 3  # Enum: MsoPictureColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecolortype
msoPictureGrayscale = 2  # Enum: MsoPictureColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecolortype
msoPictureMixed = (
    -2
)  # Enum: MsoPictureColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecolortype
msoPictureWatermark = 4  # Enum: MsoPictureColorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecolortype
msoPictureCompressDocDefault = (
    -1
)  # Enum: MsoPictureCompress, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecompress
msoPictureCompressFalse = 0  # Enum: MsoPictureCompress, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecompress
msoPictureCompressTrue = 1  # Enum: MsoPictureCompress, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturecompress
msoEffectBackgroundRemoval = 1  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectBlur = 2  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectBrightnessContrast = 3  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectCement = 4  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectChalkSketch = 6  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectColorTemperature = 7  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectCrisscrossEtching = 5  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectCutout = 8  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectFilmGrain = 9  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectGlass = 10  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectGlowDiffused = 11  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectGlowEdges = 12  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectLightScreen = 13  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectLineDrawing = 14  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectMarker = 15  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectMosiaicBubbles = 16  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectNone = 0  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPaintBrush = 17  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPaintStrokes = 18  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPastelsSmooth = 19  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPencilGrayscale = 20  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPencilSketch = 21  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPhotocopy = 22  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectPlasticWrap = 23  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectSaturation = 24  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectSharpenSoften = 25  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectTexturizer = 26  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoEffectWatercolorSponge = 27  # Enum: MsoPictureEffectType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopictureeffecttype
msoPictureTypeBMP = 1  # Enum: MsoPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturetype
msoPictureTypeDefault = (
    -2
)  # Enum: MsoPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturetype
msoPictureTypeGIF = 2  # Enum: MsoPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturetype
msoPictureTypeJPG = 3  # Enum: MsoPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturetype
msoPictureTypePDF = 4  # Enum: MsoPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturetype
msoPictureTypePNG = 0  # Enum: MsoPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopicturetype
msoCameraIsometricBottomDown = 23  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricBottomUp = 22  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricLeftDown = 25  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricLeftUp = 24  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis1Left = 28  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis1Right = 29  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis1Top = 30  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis2Left = 31  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis2Right = 32  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis2Top = 33  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis3Bottom = 36  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis3Left = 34  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis3Right = 35  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis4Bottom = 39  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis4Left = 37  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricOffAxis4Right = 38  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricRightDown = 27  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricRightUp = 26  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricTopDown = 21  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraIsometricTopUp = 20  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueBottom = 8  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueBottomLeft = 7  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueBottomRight = 9  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueFront = 5  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueLeft = 4  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueRight = 6  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueTop = 2  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueTopLeft = 1  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyObliqueTopRight = 3  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveBottom = 17  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveBottomLeft = 16  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveBottomRight = 18  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveFront = 14  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveLeft = 13  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveRight = 15  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveTop = 11  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveTopLeft = 10  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraLegacyPerspectiveTopRight = 12  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueBottom = 46  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueBottomLeft = 45  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueBottomRight = 47  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueLeft = 43  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueRight = 44  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueTop = 41  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueTopLeft = 40  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraObliqueTopRight = 42  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraOrthographicFront = 19  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveAbove = 51  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveAboveLeftFacing = 53  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveAboveRightFacing = 54  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveBelow = 52  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveContrastingLeftFacing = 55  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveContrastingRightFacing = 56  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveFront = 48  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveHeroicExtremeLeftFacing = 59  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveHeroicExtremeRightFacing = 60  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveHeroicLeftFacing = 57  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveHeroicRightFacing = 58  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveLeft = 49  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveRelaxed = 61  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveRelaxedModerately = 62  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoCameraPerspectiveRight = 50  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoPresetCameraMixed = (
    -2
)  # Enum: MsoPresetCamera, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetcamera
msoExtrusionBottom = 2  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionBottomLeft = 3  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionBottomRight = 1  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionLeft = 6  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionNone = 5  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionRight = 4  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionTop = 8  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionTopLeft = 9  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoExtrusionTopRight = 7  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoPresetExtrusionDirectionMixed = (
    -2
)  # Enum: MsoPresetExtrusionDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetextrusiondirection
msoGradientBrass = 20  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientCalmWater = 8  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientChrome = 21  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientChromeII = 22  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientDaybreak = 4  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientDesert = 6  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientEarlySunset = 1  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientFire = 9  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientFog = 10  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientGold = 18  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientGoldII = 19  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientHorizon = 5  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientLateSunset = 2  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientMahogany = 15  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientMoss = 11  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientNightfall = 3  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientOcean = 7  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientParchment = 14  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientPeacock = 12  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientRainbow = 16  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientRainbowII = 17  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientSapphire = 24  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientSilver = 23  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoGradientWheat = 13  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoPresetGradientMixed = (
    -2
)  # Enum: MsoPresetGradientType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetgradienttype
msoLightingBottom = 8  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingBottomLeft = 7  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingBottomRight = 9  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingLeft = 4  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingNone = 5  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingRight = 6  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingTop = 2  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingTopLeft = 1  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingTopRight = 3  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoPresetLightingDirectionMixed = (
    -2
)  # Enum: MsoPresetLightingDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingdirection
msoLightingBright = 3  # Enum: MsoPresetLightingSoftness, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingsoftness
msoLightingDim = 1  # Enum: MsoPresetLightingSoftness, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingsoftness
msoLightingNormal = 2  # Enum: MsoPresetLightingSoftness, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingsoftness
msoPresetLightingSoftnessMixed = (
    -2
)  # Enum: MsoPresetLightingSoftness, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetlightingsoftness
msoMaterialClear = 13  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialDarkEdge = 11  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialFlat = 14  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialMatte = 1  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialMatte2 = 5  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialMetal = 3  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialMetal2 = 7  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialPlastic = 2  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialPlastic2 = 6  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialPowder = 10  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialSoftEdge = 12  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialSoftMetal = 15  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialTranslucentPowder = 9  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialWarmMatte = 8  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoMaterialWireFrame = 4  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoPresetMaterialMixed = (
    -2
)  # Enum: MsoPresetMaterial, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetmaterial
msoTextEffect1 = 0  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect10 = 9  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect11 = 10  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect12 = 11  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect13 = 12  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect14 = 13  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect15 = 14  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect16 = 15  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect17 = 16  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect18 = 17  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect19 = 18  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect2 = 1  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect20 = 19  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect21 = 20  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect22 = 21  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect23 = 22  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect24 = 23  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect25 = 24  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect26 = 25  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect27 = 26  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect28 = 27  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect29 = 28  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect3 = 2  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect30 = 29  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect31 = 30  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect32 = 31  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect33 = 32  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect34 = 33  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect35 = 34  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect36 = 35  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect37 = 36  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect38 = 37  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect39 = 38  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect4 = 3  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect40 = 39  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect41 = 40  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect42 = 41  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect43 = 42  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect44 = 43  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect45 = 44  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect46 = 45  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect47 = 46  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect48 = 47  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect49 = 48  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect5 = 4  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect50 = 49  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect6 = 5  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect7 = 6  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect8 = 7  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffect9 = 8  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffectMixed = (
    -2
)  # Enum: MsoPresetTextEffect, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffect
msoTextEffectShapeArchDownCurve = 10  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeArchDownPour = 14  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeArchUpCurve = 9  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeArchUpPour = 13  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeButtonCurve = 12  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeButtonPour = 16  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCanDown = 20  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCanUp = 19  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCascadeDown = 40  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCascadeUp = 39  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeChevronDown = 6  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeChevronUp = 5  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCircleCurve = 11  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCirclePour = 15  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCurveDown = 18  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeCurveUp = 17  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDeflate = 26  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDeflateBottom = 28  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDeflateInflate = 31  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDeflateInflateDeflate = 32  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDeflateTop = 30  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDoubleWave1 = 23  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeDoubleWave2 = 24  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeFadeDown = 36  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeFadeLeft = 34  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeFadeRight = 33  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeFadeUp = 35  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeInflate = 25  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeInflateBottom = 27  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeInflateTop = 29  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeMixed = (
    -2
)  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapePlainText = 1  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeRingInside = 7  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeRingOutside = 8  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeSlantDown = 38  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeSlantUp = 37  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeStop = 2  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeTriangleDown = 4  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeTriangleUp = 3  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeWave1 = 21  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoTextEffectShapeWave2 = 22  # Enum: MsoPresetTextEffectShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexteffectshape
msoPresetTextureMixed = (
    -2
)  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureBlueTissuePaper = 17  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureBouquet = 20  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureBrownMarble = 11  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureCanvas = 2  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureCork = 21  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureDenim = 3  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureFishFossil = 7  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureGranite = 12  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureGreenMarble = 9  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureMediumWood = 24  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureNewsprint = 13  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureOak = 23  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTexturePaperBag = 6  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTexturePapyrus = 1  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureParchment = 15  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTexturePinkTissuePaper = 18  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTexturePurpleMesh = 19  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureRecycledPaper = 14  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureSand = 8  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureStationery = 16  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureWalnut = 22  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureWaterDroplets = 5  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureWhiteMarble = 10  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoTextureWovenMat = 4  # Enum: MsoPresetTexture, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresettexture
msoPresetThreeDFormatMixed = (
    -2
)  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD1 = 1  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD10 = 10  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD11 = 11  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD12 = 12  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD13 = 13  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD14 = 14  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD15 = 15  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD16 = 16  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD17 = 17  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD18 = 18  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD19 = 19  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD2 = 2  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD20 = 20  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD3 = 3  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD4 = 4  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD5 = 5  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD6 = 6  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD7 = 7  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD8 = 8  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoThreeD9 = 9  # Enum: MsoPresetThreeDFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msopresetthreedformat
msoReflectionType1 = 1  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType2 = 2  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType3 = 3  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType4 = 4  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType5 = 5  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType6 = 6  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType7 = 7  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType8 = 8  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionType9 = 9  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionTypeMixed = (
    -2
)  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoReflectionTypeNone = 0  # Enum: MsoReflectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoreflectiontype
msoAfterLastSibling = 4  # Enum: MsoRelativeNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msorelativenodeposition
msoAfterNode = 2  # Enum: MsoRelativeNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msorelativenodeposition
msoBeforeFirstSibling = 3  # Enum: MsoRelativeNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msorelativenodeposition
msoBeforeNode = 1  # Enum: MsoRelativeNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msorelativenodeposition
msoScaleFromBottomRight = 2  # Enum: MsoScaleFrom, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscalefrom
msoScaleFromMiddle = 1  # Enum: MsoScaleFrom, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscalefrom
msoScaleFromTopLeft = 0  # Enum: MsoScaleFrom, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscalefrom
msoScreenSize1024x768 = 4  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize1152x882 = 5  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize1152x900 = 6  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize1280x1024 = 7  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize1600x1200 = 8  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize1800x1440 = 9  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize1920x1200 = 10  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize544x376 = 0  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize640x480 = 1  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize720x512 = 2  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScreenSize800x600 = 3  # Enum: MsoScreenSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscreensize
msoScriptLanguageASP = 3  # Enum: MsoScriptLanguage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscriptlanguage
msoScriptLanguageJava = 1  # Enum: MsoScriptLanguage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscriptlanguage
msoScriptLanguageOther = 4  # Enum: MsoScriptLanguage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscriptlanguage
msoScriptLanguageVisualBasic = 2  # Enum: MsoScriptLanguage, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscriptlanguage
msoScriptLocationInBody = 2  # Enum: MsoScriptLocation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscriptlocation
msoScriptLocationInHead = 1  # Enum: MsoScriptLocation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoscriptlocation
msoSearchInCustom = 3  # Enum: MsoSearchIn, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosearchin
msoSearchInMyComputer = 0  # Enum: MsoSearchIn, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosearchin
msoSearchInMyNetworkPlaces = 2  # Enum: MsoSearchIn, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosearchin
msoSearchInOutlook = 1  # Enum: MsoSearchIn, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosearchin
msoSegmentCurve = 1  # Enum: MsoSegmentType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosegmenttype
msoSegmentLine = 0  # Enum: MsoSegmentType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosegmenttype
msoShadowStyleInnerShadow = 1  # Enum: MsoShadowStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowstyle
msoShadowStyleMixed = (
    -2
)  # Enum: MsoShadowStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowstyle
msoShadowStyleOuterShadow = 2  # Enum: MsoShadowStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowstyle
msoShadow1 = 1  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow10 = 10  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow11 = 11  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow12 = 12  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow13 = 13  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow14 = 14  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow15 = 15  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow16 = 16  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow17 = 17  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow18 = 18  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow19 = 19  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow2 = 2  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow20 = 20  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow21 = 21  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow22 = 22  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow23 = 23  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow24 = 24  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow25 = 25  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow26 = 26  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow27 = 27  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow28 = 28  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow29 = 29  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow3 = 3  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow30 = 30  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow31 = 31  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow32 = 32  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow33 = 33  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow34 = 34  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow35 = 35  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow36 = 36  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow37 = 37  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow38 = 38  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow39 = 39  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow4 = 4  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow40 = 40  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow41 = 41  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow42 = 42  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow43 = 43  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow5 = 5  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow6 = 6  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow7 = 7  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow8 = 8  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadow9 = 9  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoShadowMixed = (
    -2
)  # Enum: MsoShadowType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshadowtype
msoLineStylePreset1 = 10001  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset10 = 10010  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset11 = 10011  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset12 = 10012  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset13 = 10013  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset14 = 10014  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset15 = 10015  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset16 = 10016  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset17 = 10017  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset18 = 10018  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset19 = 10019  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset2 = 10002  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset20 = 10020  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset21 = 10021  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset22 = 10022  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset23 = 10023  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset24 = 10024  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset25 = 10025  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset26 = 10026  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset27 = 10027  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset28 = 10028  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset29 = 10029  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset3 = 10003  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset30 = 10030  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset31 = 10031  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset32 = 10032  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset33 = 10033  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset34 = 10034  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset35 = 10035  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset36 = 10036  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset37 = 10037  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset38 = 10038  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset39 = 10039  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset4 = 10004  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset40 = 10040  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset41 = 10041  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset42 = 10042  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset5 = 10005  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset6 = 10006  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset7 = 10007  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset8 = 10008  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoLineStylePreset9 = 10009  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStyleMixed = (
    -2
)  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStyleNotAPreset = 0  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset1 = 1  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset10 = 10  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset11 = 11  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset12 = 12  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset13 = 13  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset14 = 14  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset15 = 15  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset16 = 16  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset17 = 17  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset18 = 18  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset19 = 19  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset2 = 2  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset20 = 20  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset21 = 21  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset22 = 22  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset23 = 23  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset24 = 24  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset25 = 25  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset26 = 26  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset27 = 27  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset28 = 28  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset29 = 29  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset3 = 3  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset30 = 30  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset31 = 31  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset32 = 32  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset33 = 33  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset34 = 34  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset35 = 35  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset36 = 36  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset37 = 37  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset38 = 38  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset39 = 39  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset4 = 4  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset40 = 40  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset41 = 41  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset42 = 42  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset43 = 43  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset44 = 44  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset45 = 45  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset46 = 46  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset47 = 47  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset48 = 48  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset49 = 49  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset5 = 5  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset50 = 50  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset51 = 51  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset52 = 52  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset53 = 53  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset54 = 54  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset55 = 55  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset56 = 56  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset57 = 57  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset58 = 58  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset59 = 59  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset6 = 6  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset60 = 60  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset61 = 61  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset62 = 62  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset63 = 63  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset64 = 64  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset65 = 65  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset66 = 66  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset67 = 67  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset68 = 68  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset69 = 69  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset7 = 7  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset70 = 70  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset71 = 71  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset72 = 72  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset73 = 73  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset74 = 74  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset75 = 75  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset76 = 76  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset77 = 77  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset8 = 8  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
msoShapeStylePreset9 = 9  # Enum: MsoShapeStyleIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapestyleindex
mso3DModel = 30  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoAutoShape = 1  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoCallout = 2  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoCanvas = 20  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoChart = 3  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoComment = 4  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoContentApp = 27  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoDiagram = 21  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoEmbeddedOLEObject = 7  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoFormControl = 8  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoFreeform = 5  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoGraphic = 28  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoGroup = 6  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoInk = 22  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoInkComment = 23  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoLine = 9  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoLinked3DModel = 31  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoLinkedGraphic = 29  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoLinkedOLEObject = 10  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoLinkedPicture = 11  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoMedia = 16  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoOLEControlObject = 12  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoPicture = 13  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoPlaceholder = 14  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoScriptAnchor = 18  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoShapeTypeMixed = (
    -2
)  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoSlicer = 25  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoSmartArt = 24  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoTable = 19  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoTextBox = 17  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoTextEffect = 15  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoWebVideo = 26  # Enum: MsoShapeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoshapetype
msoSharedWorkspaceTaskPriorityHigh = 1  # Enum: MsoSharedWorkspaceTaskPriority, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskpriority
msoSharedWorkspaceTaskPriorityLow = 3  # Enum: MsoSharedWorkspaceTaskPriority, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskpriority
msoSharedWorkspaceTaskPriorityNormal = 2  # Enum: MsoSharedWorkspaceTaskPriority, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskpriority
msoSharedWorkspaceTaskStatusCompleted = 3  # Enum: MsoSharedWorkspaceTaskStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskstatus
msoSharedWorkspaceTaskStatusDeferred = 4  # Enum: MsoSharedWorkspaceTaskStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskstatus
msoSharedWorkspaceTaskStatusInProgress = 2  # Enum: MsoSharedWorkspaceTaskStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskstatus
msoSharedWorkspaceTaskStatusNotStarted = 1  # Enum: MsoSharedWorkspaceTaskStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskstatus
msoSharedWorkspaceTaskStatusWaiting = 5  # Enum: MsoSharedWorkspaceTaskStatus, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosharedworkspacetaskstatus
msoSignatureSubsetAll = 5  # Enum: MsoSignatureSubset, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosignaturesubset
msoSignatureSubsetSignatureLines = 2  # Enum: MsoSignatureSubset, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosignaturesubset
msoSignatureSubsetSignatureLinesSigned = 3  # Enum: MsoSignatureSubset, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosignaturesubset
msoSignatureSubsetSignatureLinesUnsigned = 4  # Enum: MsoSignatureSubset, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosignaturesubset
msoSignatureSubsetSignaturesAllSigs = 0  # Enum: MsoSignatureSubset, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosignaturesubset
msoSignatureSubsetSignaturesNonVisible = 1  # Enum: MsoSignatureSubset, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosignaturesubset
msoSmartArtNodeAbove = 4  # Enum: MsoSmartArtNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodeposition
msoSmartArtNodeAfter = 2  # Enum: MsoSmartArtNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodeposition
msoSmartArtNodeBefore = 3  # Enum: MsoSmartArtNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodeposition
msoSmartArtNodeBelow = 5  # Enum: MsoSmartArtNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodeposition
msoSmartArtNodeDefault = 1  # Enum: MsoSmartArtNodePosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodeposition
msoSmartArtNodeTypeAssistant = 2  # Enum: MsoSmartArtNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodetype
msoSmartArtNodeTypeDefault = 1  # Enum: MsoSmartArtNodeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosmartartnodetype
msoSoftEdgeType1 = 1  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeType2 = 2  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeType3 = 3  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeType4 = 4  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeType5 = 5  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeType6 = 6  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeTypeMixed = (
    -2
)  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSoftEdgeTypeNone = 0  # Enum: MsoSoftEdgeType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosoftedgetype
msoSortByFileName = 1  # Enum: MsoSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortby
msoSortByFileType = 3  # Enum: MsoSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortby
msoSortByLastModified = 4  # Enum: MsoSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortby
msoSortByNone = 5  # Enum: MsoSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortby
msoSortBySize = 2  # Enum: MsoSortBy, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortby
msoSortOrderAscending = 1  # Enum: MsoSortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortorder
msoSortOrderDescending = 2  # Enum: MsoSortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosortorder
msoSyncAvailableAnywhere = 2  # Enum: MsoSyncAvailableType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncavailabletype
msoSyncAvailableNone = 0  # Enum: MsoSyncAvailableType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncavailabletype
msoSyncAvailableOffline = 1  # Enum: MsoSyncAvailableType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncavailabletype
msoSyncCompareAndMerge = 0  # Enum: MsoSyncCompareType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynccomparetype
msoSyncCompareSideBySide = 1  # Enum: MsoSyncCompareType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynccomparetype
msoSyncConflictClientWins = 0  # Enum: MsoSyncConflictResolutionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncconflictresolutiontype
msoSyncConflictMerge = 2  # Enum: MsoSyncConflictResolutionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncconflictresolutiontype
msoSyncConflictServerWins = 1  # Enum: MsoSyncConflictResolutionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncconflictresolutiontype
msoSyncErrorCouldNotCompare = 13  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorCouldNotConnect = 2  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorCouldNotOpen = 11  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorCouldNotResolve = 14  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorCouldNotUpdate = 12  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorFileInUse = 6  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorFileNotFound = 4  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorFileTooLarge = 5  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorNoNetwork = 15  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorNone = 0  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorOutOfSpace = 3  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorUnauthorizedUser = 1  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorUnknown = 16  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorUnknownDownload = 10  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorUnknownUpload = 9  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorVirusDownload = 8  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncErrorVirusUpload = 7  # Enum: MsoSyncErrorType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncerrortype
msoSyncEventDownloadFailed = 2  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventDownloadInitiated = 0  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventDownloadNoChange = 6  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventDownloadSucceeded = 1  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventOffline = 7  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventUploadFailed = 5  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventUploadInitiated = 3  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncEventUploadSucceeded = 4  # Enum: MsoSyncEventType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosynceventtype
msoSyncStatusConflict = 4  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusError = 6  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusLatest = 1  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusLocalChanges = 3  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusNewerAvailable = 2  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusNoSharedWorkspace = 0  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusNotRoaming = 0  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncStatusSuspended = 5  # Enum: MsoSyncStatusType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncstatustype
msoSyncVersionLastViewed = 0  # Enum: MsoSyncVersionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncversiontype
msoSyncVersionServer = 1  # Enum: MsoSyncVersionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msosyncversiontype
msoTabStopCenter = 2  # Enum: MsoTabStopType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotabstoptype
msoTabStopDecimal = 4  # Enum: MsoTabStopType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotabstoptype
msoTabStopLeft = 1  # Enum: MsoTabStopType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotabstoptype
msoTabStopMixed = (
    -2
)  # Enum: MsoTabStopType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotabstoptype
msoTabStopRight = 3  # Enum: MsoTabStopType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotabstoptype
msoTargetBrowserIE4 = 2  # Enum: MsoTargetBrowser, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotargetbrowser
msoTargetBrowserIE5 = 3  # Enum: MsoTargetBrowser, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotargetbrowser
msoTargetBrowserIE6 = 4  # Enum: MsoTargetBrowser, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotargetbrowser
msoTargetBrowserV3 = 0  # Enum: MsoTargetBrowser, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotargetbrowser
msoTargetBrowserV4 = 1  # Enum: MsoTargetBrowser, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotargetbrowser
msoAllCaps = 2  # Enum: MsoTextCaps, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcaps
msoCapsMixed = (
    -2
)  # Enum: MsoTextCaps, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcaps
msoNoCaps = 0  # Enum: MsoTextCaps, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcaps
msoSmallCaps = 1  # Enum: MsoTextCaps, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcaps
msoCaseLower = 2  # Enum: MsoTextChangeCase, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextchangecase
msoCaseSentence = 1  # Enum: MsoTextChangeCase, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextchangecase
msoCaseTitle = 4  # Enum: MsoTextChangeCase, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextchangecase
msoCaseToggle = 5  # Enum: MsoTextChangeCase, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextchangecase
msoCaseUpper = 3  # Enum: MsoTextChangeCase, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextchangecase
msoCharWrapMixed = (
    -2
)  # Enum: MsoTextCharWrap, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcharwrap
msoCustomCharWrap = 3  # Enum: MsoTextCharWrap, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcharwrap
msoNoCharWrap = 0  # Enum: MsoTextCharWrap, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcharwrap
msoStandardCharWrap = 1  # Enum: MsoTextCharWrap, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcharwrap
msoStrictCharWrap = 2  # Enum: MsoTextCharWrap, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextcharwrap
msoTextDirectionLeftToRight = 1  # Enum: MsoTextDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextdirection
msoTextDirectionMixed = (
    -2
)  # Enum: MsoTextDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextdirection
msoTextDirectionRightToLeft = 2  # Enum: MsoTextDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextdirection
msoTextEffectAlignmentCentered = 2  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoTextEffectAlignmentLeft = 1  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoTextEffectAlignmentLetterJustify = 4  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoTextEffectAlignmentMixed = (
    -2
)  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoTextEffectAlignmentRight = 3  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoTextEffectAlignmentStretchJustify = 6  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoTextEffectAlignmentWordJustify = 5  # Enum: MsoTextEffectAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexteffectalignment
msoFontAlignAuto = 0  # Enum: MsoTextFontAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextfontalign
msoFontAlignBaseline = 3  # Enum: MsoTextFontAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextfontalign
msoFontAlignBottom = 4  # Enum: MsoTextFontAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextfontalign
msoFontAlignCenter = 2  # Enum: MsoTextFontAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextfontalign
msoFontAlignMixed = (
    -2
)  # Enum: MsoTextFontAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextfontalign
msoFontAlignTop = 1  # Enum: MsoTextFontAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextfontalign
msoTextOrientationDownward = 3  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoTextOrientationHorizontal = 1  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoTextOrientationHorizontalRotatedFarEast = 6  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoTextOrientationMixed = (
    -2
)  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoTextOrientationUpward = 2  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoTextOrientationVertical = 5  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoTextOrientationVerticalFarEast = 4  # Enum: MsoTextOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextorientation
msoMsoTextRangeInsertAfter = 1  # Enum: MsoTextRangeInsertPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextrangeinsertposition
msoMsoTextRangeInsertBefore = 0  # Enum: MsoTextRangeInsertPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextrangeinsertposition
msoDoubleStrike = 2  # Enum: MsoTextStrike, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextstrike
msoNoStrike = 0  # Enum: MsoTextStrike, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextstrike
msoSingleStrike = 1  # Enum: MsoTextStrike, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextstrike
msoStrikeMixed = (
    -2
)  # Enum: MsoTextStrike, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextstrike
msoTabAlignCenter = 1  # Enum: MsoTextTabAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexttabalign
msoTabAlignDecimal = 3  # Enum: MsoTextTabAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexttabalign
msoTabAlignLeft = 0  # Enum: MsoTextTabAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexttabalign
msoTabAlignMixed = (
    -2
)  # Enum: MsoTextTabAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexttabalign
msoTabAlignRight = 2  # Enum: MsoTextTabAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexttabalign
msoNoUnderline = 0  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDashHeavyLine = 8  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDashLine = 7  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDashLongHeavyLine = 10  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDashLongLine = 9  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDotDashHeavyLine = 12  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDotDashLine = 11  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDotDotDashHeavyLine = 14  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDotDotDashLine = 13  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDottedHeavyLine = 6  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDottedLine = 5  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineDoubleLine = 3  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineHeavyLine = 4  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineMixed = (
    -2
)  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineSingleLine = 2  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineWavyDoubleLine = 17  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineWavyHeavyLine = 16  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineWavyLine = 15  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoUnderlineWords = 1  # Enum: MsoTextUnderlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotextunderlinetype
msoTextureAlignmentMixed = (
    -2
)  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureBottom = 7  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureBottomLeft = 6  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureBottomRight = 8  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureCenter = 4  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureLeft = 3  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureRight = 5  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureTop = 1  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureTopLeft = 0  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTextureTopRight = 2  # Enum: MsoTextureAlignment, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturealignment
msoTexturePreset = 1  # Enum: MsoTextureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturetype
msoTextureTypeMixed = (
    -2
)  # Enum: MsoTextureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturetype
msoTextureUserDefined = 2  # Enum: MsoTextureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotexturetype
msoNotThemeColor = 0  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorAccent1 = 5  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorAccent2 = 6  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorAccent3 = 7  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorAccent4 = 8  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorAccent5 = 9  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorAccent6 = 10  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorBackground1 = 14  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorBackground2 = 16  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorDark1 = 1  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorDark2 = 3  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorFollowedHyperlink = 12  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorHyperlink = 11  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorLight1 = 2  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorLight2 = 4  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorMixed = (
    -2
)  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorText1 = 13  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeColorText2 = 15  # Enum: MsoThemeColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorindex
msoThemeAccent1 = 5  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeAccent2 = 6  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeAccent3 = 7  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeAccent4 = 8  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeAccent5 = 9  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeAccent6 = 10  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeDark1 = 1  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeDark2 = 3  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeFollowedHyperlink = 12  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeHyperlink = 11  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeLight1 = 2  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoThemeLight2 = 4  # Enum: MsoThemeColorSchemeIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.msothemecolorschemeindex
msoCTrue = 1  # Enum: MsoTriState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotristate
msoFalse = 0  # Enum: MsoTriState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotristate
msoTriStateMixed = (
    -2
)  # Enum: MsoTriState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotristate
msoTriStateToggle = (
    -3
)  # Enum: MsoTriState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotristate
msoTrue = (
    -1
)  # Enum: MsoTriState, https://docs.microsoft.com/ja-jp/office/vba/api/office.msotristate
msoAnchorBottom = 4  # Enum: MsoVerticalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoverticalanchor
msoAnchorBottomBaseLine = 5  # Enum: MsoVerticalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoverticalanchor
msoAnchorMiddle = 3  # Enum: MsoVerticalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoverticalanchor
msoAnchorTop = 1  # Enum: MsoVerticalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoverticalanchor
msoAnchorTopBaseline = 2  # Enum: MsoVerticalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoverticalanchor
msoVerticalAnchorMixed = (
    -2
)  # Enum: MsoVerticalAnchor, https://docs.microsoft.com/ja-jp/office/vba/api/office.msoverticalanchor
msoWarpFormat1 = 0  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat10 = 9  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat11 = 10  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat12 = 11  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat13 = 12  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat14 = 13  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat15 = 14  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat16 = 15  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat17 = 16  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat18 = 17  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat19 = 18  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat2 = 1  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat20 = 19  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat21 = 20  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat22 = 21  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat23 = 22  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat24 = 23  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat25 = 24  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat26 = 25  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat27 = 26  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat28 = 27  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat29 = 28  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat3 = 2  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat30 = 29  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat31 = 30  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat32 = 31  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat33 = 32  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat34 = 33  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat35 = 34  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat36 = 35  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat37 = 36  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat4 = 3  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat5 = 4  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat6 = 5  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat7 = 6  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat8 = 7  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormat9 = 8  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWarpFormatMixed = (
    -2
)  # Enum: MsoWarpFormat, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowarpformat
msoWizardActActive = 1  # Enum: MsoWizardActType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardacttype
msoWizardActInactive = 0  # Enum: MsoWizardActType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardacttype
msoWizardActResume = 3  # Enum: MsoWizardActType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardacttype
msoWizardActSuspend = 2  # Enum: MsoWizardActType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardacttype
msoWizardMsgLocalStateOff = 2  # Enum: MsoWizardMsgType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardmsgtype
msoWizardMsgLocalStateOn = 1  # Enum: MsoWizardMsgType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardmsgtype
msoWizardMsgResuming = 5  # Enum: MsoWizardMsgType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardmsgtype
msoWizardMsgShowHelp = 3  # Enum: MsoWizardMsgType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardmsgtype
msoWizardMsgSuspending = 4  # Enum: MsoWizardMsgType, https://docs.microsoft.com/ja-jp/office/vba/api/office.msowizardmsgtype
msoBringForward = 2  # Enum: MsoZOrderCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msozordercmd
msoBringInFrontOfText = 4  # Enum: MsoZOrderCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msozordercmd
msoBringToFront = 0  # Enum: MsoZOrderCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msozordercmd
msoSendBackward = 3  # Enum: MsoZOrderCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msozordercmd
msoSendBehindText = 5  # Enum: MsoZOrderCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msozordercmd
msoSendToBack = 1  # Enum: MsoZOrderCmd, https://docs.microsoft.com/ja-jp/office/vba/api/office.msozordercmd
RibbonControlSizeLarge = 1  # Enum: RibbonControlSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.ribboncontrolsize
RibbonControlSizeRegular = 0  # Enum: RibbonControlSize, https://docs.microsoft.com/ja-jp/office/vba/api/office.ribboncontrolsize
sigdetApplicationName = 1  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetApplicationVersion = 2  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetColorDepth = 8  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDelSuggSigner = 16  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDelSuggSignerEmail = 20  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDelSuggSignerEmailSet = 21  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDelSuggSignerLine2 = 18  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDelSuggSignerLine2Set = 19  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDelSuggSignerSet = 17  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetDocPreviewImg = 10  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetHashAlgorithm = 14  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetHorizResolution = 6  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetIPCurrentView = 12  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetIPFormHash = 11  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetLocalSigningTime = 0  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetNumberOfMonitors = 5  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetOfficeVersion = 3  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetShouldShowViewWarning = 15  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetSignatureType = 13  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetSignedData = 9  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetVertResolution = 7  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
sigdetWindowsVersion = 4  # Enum: SignatureDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturedetail
siglnimgSigned = 4  # Enum: SignatureLineImage, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturelineimage
siglnimgSignedInvalid = 3  # Enum: SignatureLineImage, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturelineimage
siglnimgSignedValid = 2  # Enum: SignatureLineImage, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturelineimage
siglnimgSoftwareRequired = 0  # Enum: SignatureLineImage, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturelineimage
siglnimgUnsigned = 1  # Enum: SignatureLineImage, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturelineimage
sigprovdetHashAlgorithm = 1  # Enum: SignatureProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signatureproviderdetail
sigprovdetUIOnly = 2  # Enum: SignatureProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signatureproviderdetail
sigprovdetUrl = 0  # Enum: SignatureProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signatureproviderdetail
sigprovdetUseOfficeStampUI = 4  # Enum: SignatureProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signatureproviderdetail
sigprovdetUseOfficeUI = 3  # Enum: SignatureProviderDetail, https://docs.microsoft.com/ja-jp/office/vba/api/office.signatureproviderdetail
sigtypeMax = 3  # Enum: SignatureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturetype
sigtypeNonVisible = 1  # Enum: SignatureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturetype
sigtypeSignatureLine = 2  # Enum: SignatureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturetype
sigtypeUnknown = 0  # Enum: SignatureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.signaturetype
xlAxisCrossesAutomatic = (
    -4105
)  # Enum: XlAxisCrosses, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxiscrosses
xlAxisCrossesCustom = (
    -4114
)  # Enum: XlAxisCrosses, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxiscrosses
xlAxisCrossesMaximum = 2  # Enum: XlAxisCrosses, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxiscrosses
xlAxisCrossesMinimum = 4  # Enum: XlAxisCrosses, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxiscrosses
xlPrimary = 1  # Enum: XlAxisGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxisgroup
xlSecondary = 2  # Enum: XlAxisGroup, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxisgroup
xlCategory = 1  # Enum: XlAxisType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxistype
xlSeriesAxis = 3  # Enum: XlAxisType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxistype
xlValue = 2  # Enum: XlAxisType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlaxistype
xlBox = 0  # Enum: XlBarShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbarshape
xlConeToMax = 5  # Enum: XlBarShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbarshape
xlConeToPoint = 4  # Enum: XlBarShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbarshape
xlCylinder = 3  # Enum: XlBarShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbarshape
xlPyramidToMax = 2  # Enum: XlBarShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbarshape
xlPyramidToPoint = 1  # Enum: XlBarShape, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbarshape
xlBinsTypeAutomatic = 0  # Enum: XlBinsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbinstype
xlBinsTypeBinCount = 4  # Enum: XlBinsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbinstype
xlBinsTypeBinSize = 3  # Enum: XlBinsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbinstype
xlBinsTypeCategorical = 1  # Enum: XlBinsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbinstype
xlBinsTypeManual = 2  # Enum: XlBinsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlbinstype
xlHairline = 1  # Enum: XlBorderWeight, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlborderweight
xlMedium = (
    -4138
)  # Enum: XlBorderWeight, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlborderweight
xlThick = 4  # Enum: XlBorderWeight, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlborderweight
xlThin = 2  # Enum: XlBorderWeight, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlborderweight
xlCategoryLabelLevelAll = (
    -1
)  # Enum: XlCategoryLabelLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorylabellevel
xlCategoryLabelLevelCustom = (
    -2
)  # Enum: XlCategoryLabelLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorylabellevel
xlCategoryLabelLevelNone = (
    -3
)  # Enum: XlCategoryLabelLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorylabellevel
xlCategoryAscending = 2  # Enum: XlCategorySortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorysortorder
xlCategoryDescending = 3  # Enum: XlCategorySortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorysortorder
xlIndexAscending = 0  # Enum: XlCategorySortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorysortorder
xlIndexDescending = 1  # Enum: XlCategorySortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorysortorder
xlAutomaticScale = (
    -4105
)  # Enum: XlCategoryType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorytype
xlCategoryScale = 2  # Enum: XlCategoryType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorytype
xlTimeScale = 3  # Enum: XlCategoryType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcategorytype
xlChartElementPositionAutomatic = (
    -4105
)  # Enum: XlChartElementPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartelementposition
xlChartElementPositionCustom = (
    -4114
)  # Enum: XlChartElementPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartelementposition
xlAxis = 21  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlAxisTitle = 17  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlChartArea = 2  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlChartTitle = 4  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlCorners = 6  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlDataLabel = 0  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlDataTable = 7  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlDisplayUnitLabel = 30  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlDownBars = 20  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlDropLines = 26  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlErrorBars = 9  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlFloor = 23  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlHiLoLines = 25  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlLeaderLines = 29  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlLegend = 24  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlLegendEntry = 12  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlLegendKey = 13  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlMajorGridlines = 15  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlMinorGridlines = 16  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlNothing = 28  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlPivotChartCollapseEntireFieldButton = 34  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlPivotChartDropZone = 32  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlPivotChartExpandEntireFieldButton = 33  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlPivotChartFieldButton = 31  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlPlotArea = 19  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlRadarAxisLabels = 27  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlSeries = 3  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlSeriesLines = 22  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlShape = 14  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlTrendline = 8  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlUpBars = 18  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlWalls = 5  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlXErrorBars = 10  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlYErrorBars = 11  # Enum: XlChartItem, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartitem
xlDownward = (
    -4170
)  # Enum: XlChartOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartorientation
xlHorizontal = (
    -4128
)  # Enum: XlChartOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartorientation
xlUpward = (
    -4171
)  # Enum: XlChartOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartorientation
xlVertical = (
    -4166
)  # Enum: XlChartOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartorientation
xlStack = 2  # Enum: XlChartPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartpicturetype
xlStackScale = 3  # Enum: XlChartPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartpicturetype
xlStretch = 1  # Enum: XlChartPictureType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartpicturetype
xlSplitByCustomSplit = 4  # Enum: XlChartSplitType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartsplittype
xlSplitByPercentValue = 3  # Enum: XlChartSplitType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartsplittype
xlSplitByPosition = 1  # Enum: XlChartSplitType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartsplittype
xlSplitByValue = 2  # Enum: XlChartSplitType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlchartsplittype
xl3DArea = (
    -4098
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DAreaStacked = 78  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DAreaStacked100 = 79  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DBarClustered = 60  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DBarStacked = 61  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DBarStacked100 = 62  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DColumn = (
    -4100
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DColumnClustered = 54  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DColumnStacked = 55  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DColumnStacked100 = 56  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DLine = (
    -4101
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DPie = (
    -4102
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xl3DPieExploded = 70  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlArea = 1  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlAreaEx = 135  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlAreaStacked = 76  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlAreaStacked100 = 77  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlAreaStacked100Ex = 137  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlAreaStackedEx = 136  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarClustered = 57  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarClusteredEx = 132  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarOfPie = 71  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarStacked = 58  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarStacked100 = 59  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarStacked100Ex = 134  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBarStackedEx = 133  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBoxwhisker = 121  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBubble = 15  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBubble3DEffect = 87  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlBubbleEx = 139  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColumnClustered = 51  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColumnClusteredEx = 124  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColumnStacked = 52  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColumnStacked100 = 53  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColumnStacked100Ex = 126  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColumnStackedEx = 125  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCombo = (
    -4152
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlComboAreaStackedColumnClustered = 115  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlComboColumnClusteredLine = 113  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlComboColumnClusteredLineSecondaryAxis = 114  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeBarClustered = 102  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeBarStacked = 103  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeBarStacked100 = 104  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeCol = 105  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeColClustered = 99  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeColStacked = 100  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlConeColStacked100 = 101  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderBarClustered = 95  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderBarStacked = 96  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderBarStacked100 = 97  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderCol = 98  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderColClustered = 92  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderColStacked = 93  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlCylinderColStacked100 = 94  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlDoughnut = (
    -4120
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlDoughnutEx = 131  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlDoughnutExploded = 80  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlFunnel = 123  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlHistogram = 118  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLine = 4  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineEx = 127  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineMarkers = 65  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineMarkersStacked = 66  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineMarkersStacked100 = 67  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineStacked = 63  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineStacked100 = 64  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineStacked100Ex = 129  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlLineStackedEx = 128  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlOtherCombinations = 116  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPareto = 122  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPie = 5  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPieEx = 130  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPieExploded = 69  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPieOfPie = 68  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidBarClustered = 109  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidBarStacked = 110  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidBarStacked100 = 111  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidCol = 112  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidColClustered = 106  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidColStacked = 107  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlPyramidColStacked100 = 108  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlRadar = (
    -4151
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlRadarFilled = 82  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlRadarMarkers = 81  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlRegionMap = 140  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlStockHLC = 88  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlStockOHLC = 89  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlStockVHLC = 90  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlStockVOHLC = 91  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlSuggestedChart = (
    -2
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlSunburst = 120  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlSurface = 83  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlSurfaceTopView = 85  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlSurfaceTopViewWireframe = 86  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlSurfaceWireframe = 84  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlTreemap = 117  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlWaterfall = 119  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlXYScatter = (
    -4169
)  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlXYScatterEx = 138  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlXYScatterLines = 74  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlXYScatterLinesNoMarkers = 75  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlXYScatterSmooth = 72  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlXYScatterSmoothNoMarkers = 73  # Enum: XlChartType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcharttype
xlColorIndexAutomatic = (
    -4105
)  # Enum: XlColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcolorindex
xlColorIndexNone = (
    -4142
)  # Enum: XlColorIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlcolorindex
xl3DBar = (
    -4099
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xl3DSurface = (
    -4103
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlAbove = 0  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlAutomatic = (
    -4105
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlBar = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlBelow = 1  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlBoth = 1  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlBottom = (
    -4107
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCenter = (
    -4108
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlChecker = 9  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCircle = 8  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlColumn = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCombination = (
    -4111
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCorner = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCrissCross = 16  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCross = 4  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlCustom = (
    -4114
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlDefaultAutoFormat = (
    -1
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlDiamond = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlDistributed = (
    -4117
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlFill = 5  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlFixedValue = 1  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGeneral = 1  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGray16 = 17  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGray25 = (
    -4124
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGray50 = (
    -4125
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGray75 = (
    -4126
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGray8 = 18  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlGrid = 15  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlHigh = (
    -4127
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlInside = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlJustify = (
    -4130
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLeft = (
    -4131
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLightDown = 13  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLightHorizontal = 11  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLightUp = 14  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLightVertical = 12  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLow = (
    -4134
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlMaximum = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlMinimum = 4  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlMinusValues = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlNextToAxis = 4  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlNone = (
    -4142
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlOpaque = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlOutside = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlPercent = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlPlus = 9  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlPlusValues = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlRight = (
    -4152
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlScale = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlSemiGray75 = 10  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlShowLabel = 4  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlShowLabelAndPercent = 5  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlShowPercent = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlShowValue = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlSingle = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlSolid = 1  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlSquare = 1  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlStError = 4  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlStar = 5  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlTop = (
    -4160
)  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlTransparent = 2  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlTriangle = 3  # Enum: XlConstants, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlconstants
xlLabelPositionAbove = 0  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionBelow = 1  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionBestFit = 5  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionCenter = (
    -4108
)  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionCustom = 7  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionInsideBase = 4  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionInsideEnd = 3  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionLeft = (
    -4131
)  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionMixed = 6  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionOutsideEnd = 2  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlLabelPositionRight = (
    -4152
)  # Enum: XlDataLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelposition
xlDataLabelsShowBubbleSizes = 6  # Enum: XlDataLabelsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelstype
xlDataLabelsShowLabel = 4  # Enum: XlDataLabelsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelstype
xlDataLabelsShowLabelAndPercent = 5  # Enum: XlDataLabelsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelstype
xlDataLabelsShowNone = (
    -4142
)  # Enum: XlDataLabelsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelstype
xlDataLabelsShowPercent = 3  # Enum: XlDataLabelsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelstype
xlDataLabelsShowValue = 2  # Enum: XlDataLabelsType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldatalabelstype
xlInterpolated = 3  # Enum: XlDisplayBlanksAs, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayblanksas
xlNotPlotted = 1  # Enum: XlDisplayBlanksAs, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayblanksas
xlZero = 2  # Enum: XlDisplayBlanksAs, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayblanksas
xlDisplayUnitCustom = (
    -4114
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlDisplayUnitNone = (
    -4142
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlHundredMillions = (
    -8
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlHundredThousands = (
    -5
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlHundreds = (
    -2
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlMillionMillions = (
    -10
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlMillions = (
    -6
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlTenMillions = (
    -7
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlTenThousands = (
    -4
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlThousandMillions = (
    -9
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlThousands = (
    -3
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xldisplayunit
xlCap = 1  # Enum: XlEndStyleCap, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlendstylecap
xlNoCap = 2  # Enum: XlEndStyleCap, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlendstylecap
xlChartX = (
    -4168
)  # Enum: XlErrorBarDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbardirection
xlChartY = 1  # Enum: XlErrorBarDirection, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbardirection
xlErrorBarIncludeBoth = 1  # Enum: XlErrorBarInclude, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbarinclude
xlErrorBarIncludeMinusValues = 3  # Enum: XlErrorBarInclude, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbarinclude
xlErrorBarIncludeNone = (
    -4142
)  # Enum: XlErrorBarInclude, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbarinclude
xlErrorBarIncludePlusValues = 2  # Enum: XlErrorBarInclude, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbarinclude
xlErrorBarTypeCustom = (
    -4114
)  # Enum: XlErrorBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbartype
xlErrorBarTypeFixedValue = 1  # Enum: XlErrorBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbartype
xlErrorBarTypePercent = 2  # Enum: XlErrorBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbartype
xlErrorBarTypeStDev = (
    -4155
)  # Enum: XlErrorBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbartype
xlErrorBarTypeStError = 4  # Enum: XlErrorBarType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlerrorbartype
xlGeoMappingLevelAutomatic = 0  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelCountryRegion = 5  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelCountryRegionList = 6  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelCounty = 3  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelDataOnly = 1  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelPostalCode = 2  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelState = 4  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoMappingLevelWorld = 7  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeomappinglevel
xlGeoProjectionTypeAlbers = 3  # Enum: XlGeoProjectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeoprojectiontype
xlGeoProjectionTypeAutomatic = 0  # Enum: XlGeoProjectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeoprojectiontype
xlGeoProjectionTypeMercator = 1  # Enum: XlGeoProjectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeoprojectiontype
xlGeoProjectionTypeMiller = 2  # Enum: XlGeoProjectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeoprojectiontype
xlGeoProjectionTypeRobinson = 4  # Enum: XlGeoProjectionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgeoprojectiontype
xlGradientStopPositionTypeExtremeValue = 0  # Enum: XlGradientStopPositionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgradientstoppositiontype
xlGradientStopPositionTypeNumber = 1  # Enum: XlGradientStopPositionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgradientstoppositiontype
xlGradientStopPositionTypePercent = 2  # Enum: XlGradientStopPositionType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlgradientstoppositiontype
xlHAlignCenter = (
    -4108
)  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
xlHAlignCenterAcrossSelection = (
    7  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
)
xlHAlignDistributed = (
    -4117
)  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
xlHAlignFill = (
    5  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
)
xlHAlignGeneral = (
    1  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
)
xlHAlignJustify = (
    -4130
)  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
xlHAlignLeft = (
    -4131
)  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
xlHAlignRight = (
    -4152
)  # Enum: XlHAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlhalign
xlLegendPositionBottom = (
    -4107
)  # Enum: XlLegendPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xllegendposition
xlLegendPositionCorner = 2  # Enum: XlLegendPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xllegendposition
xlLegendPositionCustom = (
    -4161
)  # Enum: XlLegendPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xllegendposition
xlLegendPositionLeft = (
    -4131
)  # Enum: XlLegendPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xllegendposition
xlLegendPositionRight = (
    -4152
)  # Enum: XlLegendPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xllegendposition
xlLegendPositionTop = (
    -4160
)  # Enum: XlLegendPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xllegendposition
xlMarkerStyleAutomatic = (
    -4105
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleCircle = 8  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleDash = (
    -4115
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleDiamond = 2  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleDot = (
    -4118
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleNone = (
    -4142
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStylePicture = (
    -4147
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStylePlus = 9  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleSquare = 1  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleStar = 5  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleTriangle = 3  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlMarkerStyleX = (
    -4168
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlmarkerstyle
xlParentDataLabelOptionsBanner = 1  # Enum: XlParentDataLabelOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlparentdatalabeloptions
xlParentDataLabelOptionsNone = 0  # Enum: XlParentDataLabelOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlparentdatalabeloptions
xlParentDataLabelOptionsOverlapping = 2  # Enum: XlParentDataLabelOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlparentdatalabeloptions
xlCenterPoint = 5  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlInnerCenterPoint = 8  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlInnerClockwisePoint = 7  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlInnerCounterClockwisePoint = 9  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlMidClockwiseRadiusPoint = 4  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlMidCounterClockwiseRadiusPoint = 6  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlOuterCenterPoint = 2  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlOuterClockwisePoint = 3  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlOuterCounterClockwisePoint = 1  # Enum: XlPieSliceIndex, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpiesliceindex
xlHorizontalCoordinate = 1  # Enum: XlPieSliceLocation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpieslicelocation
xlVerticalCoordinate = 2  # Enum: XlPieSliceLocation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpieslicelocation
xlColumnField = 2  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpivotfieldorientation
xlDataField = 4  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpivotfieldorientation
xlHidden = 0  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpivotfieldorientation
xlPageField = 3  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpivotfieldorientation
xlRowField = 1  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlpivotfieldorientation
xlContext = (
    -5002
)  # Enum: XlReadingOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlreadingorder
xlLTR = (
    -5003
)  # Enum: XlReadingOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlreadingorder
xlRTL = (
    -5004
)  # Enum: XlReadingOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlreadingorder
xlRegionLabelOptionsBestFitOnly = 1  # Enum: XlRegionLabelOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlregionlabeloptions
xlRegionLabelOptionsNone = 0  # Enum: XlRegionLabelOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlregionlabeloptions
xlRegionLabelOptionsShowAll = 2  # Enum: XlRegionLabelOptions, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlregionlabeloptions
xlColumns = (
    2  # Enum: XlRowCol, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlrowcol
)
xlRows = (
    1  # Enum: XlRowCol, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlrowcol
)
xlScaleLinear = (
    -4132
)  # Enum: XlScaleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlscaletype
xlScaleLogarithmic = (
    -4133
)  # Enum: XlScaleType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlscaletype
xlSeriesColorGradientStyleDiverging = 1  # Enum: XlSeriesColorGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlseriescolorgradientstyle
xlSeriesColorGradientStyleSequential = 0  # Enum: XlSeriesColorGradientStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlseriescolorgradientstyle
xlSeriesNameLevelAll = (
    -1
)  # Enum: XlSeriesNameLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlseriesnamelevel
xlSeriesNameLevelCustom = (
    -2
)  # Enum: XlSeriesNameLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlseriesnamelevel
xlSeriesNameLevelNone = (
    -3
)  # Enum: XlSeriesNameLevel, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlseriesnamelevel
xlSizeIsArea = 1  # Enum: XlSizeRepresents, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlsizerepresents
xlSizeIsWidth = 2  # Enum: XlSizeRepresents, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlsizerepresents
xlTickLabelOrientationAutomatic = (
    -4105
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelorientation
xlTickLabelOrientationDownward = (
    -4170
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelorientation
xlTickLabelOrientationHorizontal = (
    -4128
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelorientation
xlTickLabelOrientationUpward = (
    -4171
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelorientation
xlTickLabelOrientationVertical = (
    -4166
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelorientation
xlTickLabelPositionHigh = (
    -4127
)  # Enum: XlTickLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelposition
xlTickLabelPositionLow = (
    -4134
)  # Enum: XlTickLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelposition
xlTickLabelPositionNextToAxis = 4  # Enum: XlTickLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelposition
xlTickLabelPositionNone = (
    -4142
)  # Enum: XlTickLabelPosition, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlticklabelposition
xlTickMarkCross = 4  # Enum: XlTickMark, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltickmark
xlTickMarkInside = 2  # Enum: XlTickMark, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltickmark
xlTickMarkNone = (
    -4142
)  # Enum: XlTickMark, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltickmark
xlTickMarkOutside = 3  # Enum: XlTickMark, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltickmark
xlDays = 0  # Enum: XlTimeUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltimeunit
xlMonths = 1  # Enum: XlTimeUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltimeunit
xlYears = 2  # Enum: XlTimeUnit, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltimeunit
xlExponential = 5  # Enum: XlTrendlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltrendlinetype
xlLinear = (
    -4132
)  # Enum: XlTrendlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltrendlinetype
xlLogarithmic = (
    -4133
)  # Enum: XlTrendlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltrendlinetype
xlMovingAvg = 6  # Enum: XlTrendlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltrendlinetype
xlPolynomial = 3  # Enum: XlTrendlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltrendlinetype
xlPower = 4  # Enum: XlTrendlineType, https://docs.microsoft.com/ja-jp/office/vba/api/office.xltrendlinetype
xlUnderlineStyleDouble = (
    -4119
)  # Enum: XlUnderlineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlunderlinestyle
xlUnderlineStyleDoubleAccounting = 5  # Enum: XlUnderlineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlunderlinestyle
xlUnderlineStyleNone = (
    -4142
)  # Enum: XlUnderlineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlunderlinestyle
xlUnderlineStyleSingle = 2  # Enum: XlUnderlineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlunderlinestyle
xlUnderlineStyleSingleAccounting = 4  # Enum: XlUnderlineStyle, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlunderlinestyle
xlVAlignBottom = (
    -4107
)  # Enum: XlVAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvalign
xlVAlignCenter = (
    -4108
)  # Enum: XlVAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvalign
xlVAlignDistributed = (
    -4117
)  # Enum: XlVAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvalign
xlVAlignJustify = (
    -4130
)  # Enum: XlVAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvalign
xlVAlignTop = (
    -4160
)  # Enum: XlVAlign, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvalign
xlValueAscending = 1  # Enum: XlValueSortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvaluesortorder
xlValueDescending = 2  # Enum: XlValueSortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvaluesortorder
xlValueNone = 0  # Enum: XlValueSortOrder, https://docs.microsoft.com/ja-jp/office/vba/api/office.xlvaluesortorder
msoAnimAccumulateAlways = 2  # Enum: MsoAnimAccumulate, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaccumulate
msoAnimAccumulateNone = 1  # Enum: MsoAnimAccumulate, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaccumulate
msoAnimAdditiveAddBase = 1  # Enum: MsoAnimAdditive, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimadditive
msoAnimAdditiveAddSum = 2  # Enum: MsoAnimAdditive, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimadditive
msoAnimAfterEffectDim = 1  # Enum: MsoAnimAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaftereffect
msoAnimAfterEffectHide = 2  # Enum: MsoAnimAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaftereffect
msoAnimAfterEffectHideOnNextClick = 3  # Enum: MsoAnimAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaftereffect
msoAnimAfterEffectMixed = (
    -1
)  # Enum: MsoAnimAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaftereffect
msoAnimAfterEffectNone = 0  # Enum: MsoAnimAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimaftereffect
msoAnimCommandTypeCall = 1  # Enum: MsoAnimCommandType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimcommandtype
msoAnimCommandTypeEvent = 0  # Enum: MsoAnimCommandType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimcommandtype
msoAnimCommandTypeVerb = 2  # Enum: MsoAnimCommandType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimcommandtype
msoAnimDirectionAcross = 18  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionBottom = 11  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionBottomLeft = 15  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionBottomRight = 14  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionCenter = 28  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionClockwise = 21  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionCounterclockwise = 22  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionCycleClockwise = 43  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionCycleCounterclockwise = 44  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionDown = 3  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionDownLeft = 9  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionDownRight = 8  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionFontAllCaps = 40  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionFontBold = 35  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionFontItalic = 36  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionFontShadow = 39  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionFontStrikethrough = 38  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionFontUnderline = 37  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionGradual = 42  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionHorizontal = 16  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionHorizontalIn = 23  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionHorizontalOut = 24  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionIn = 19  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionInBottom = 31  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionInCenter = 30  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionInSlightly = 29  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionInstant = 41  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionLeft = 4  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionNone = 0  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionOrdinalMask = 5  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionOut = 20  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionOutBottom = 34  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionOutCenter = 33  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionOutSlightly = 32  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionRight = 2  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionSlightly = 27  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionTop = 10  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionTopLeft = 12  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionTopRight = 13  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionUp = 1  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionUpLeft = 6  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionUpRight = 7  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionVertical = 17  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionVerticalIn = 25  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimDirectionVerticalOut = 26  # Enum: MsoAnimDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimdirection
msoAnimEffect3DArrive = 151  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffect3DJumpAndTurn = 154  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffect3DSwing = 153  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffect3DTurntable = 152  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectAppear = 1  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectArcUp = 47  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectAscend = 39  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBlast = 64  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBlinds = 3  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBoldFlash = 63  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBoldReveal = 65  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBoomerang = 25  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBounce = 26  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBox = 4  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBrushOnColor = 66  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectBrushOnUnderline = 67  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectCenterRevolve = 40  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectChangeFillColor = 54  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectChangeFont = 55  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectChangeFontColor = 56  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectChangeFontSize = 57  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectChangeFontStyle = 58  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectChangeLineColor = 60  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectCheckerboard = 5  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectCircle = 6  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectColorBlend = 68  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectColorReveal = 27  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectColorWave = 69  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectComplementaryColor = 70  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectComplementaryColor2 = 71  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectContrastingColor = 72  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectCrawl = 7  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectCredits = 28  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectCustom = 0  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectDarken = 73  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectDesaturate = 74  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectDescend = 42  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectDiamond = 8  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectDissolve = 9  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectEaseIn = 29  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectExpand = 50  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFade = 10  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFadedSwivel = 41  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFadedZoom = 48  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFlashBulb = 75  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFlashOnce = 11  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFlicker = 76  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFlip = 51  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFloat = 30  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFly = 2  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectFold = 53  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectGlide = 49  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectGrowAndTurn = 31  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectGrowShrink = 59  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectGrowWithColor = 77  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectLightSpeed = 32  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectLighten = 78  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectMediaPause = 84  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectMediaPlay = 83  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectMediaPlayFromBookmark = 150  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectMediaStop = 85  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPath4PointStar = 101  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPath5PointStar = 90  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPath6PointStar = 96  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPath8PointStar = 102  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathArcDown = 122  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathArcLeft = 136  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathArcRight = 143  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathArcUp = 129  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathBean = 116  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathBounceLeft = 126  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathBounceRight = 139  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathBuzzsaw = 110  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCircle = 86  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCrescentMoon = 91  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCurvedSquare = 105  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCurvedX = 106  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCurvyLeft = 133  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCurvyRight = 146  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathCurvyStar = 108  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathDecayingWave = 145  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathDiagonalDownRight = 134  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathDiagonalUpRight = 141  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathDiamond = 88  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathDown = 127  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathEqualTriangle = 98  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathFigure8Four = 113  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathFootball = 97  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathFunnel = 137  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathHeart = 94  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathHeartbeat = 130  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathHexagon = 89  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathHorizontalFigure8 = 111  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathInvertedSquare = 119  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathInvertedTriangle = 118  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathLeft = 120  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathLoopdeLoop = 109  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathNeutron = 114  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathOctagon = 95  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathParallelogram = 99  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathPeanut = 112  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathPentagon = 100  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathPlus = 117  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathPointyStar = 104  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathRight = 149  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathRightTriangle = 87  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSCurve1 = 144  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSCurve2 = 124  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSineWave = 125  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSpiralLeft = 140  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSpiralRight = 131  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSpring = 138  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSquare = 92  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathStairsDown = 147  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathSwoosh = 115  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathTeardrop = 103  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathTrapezoid = 93  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathTurnDown = 135  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathTurnRight = 121  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathTurnUp = 128  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathTurnUpRight = 142  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathUp = 148  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathVerticalFigure8 = 107  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathWave = 132  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPathZigzag = 123  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPeek = 12  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPinwheel = 33  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectPlus = 13  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectRandomBars = 14  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectRandomEffects = 24  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectRiseUp = 34  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectShimmer = 52  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSling = 43  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSpin = 61  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSpinner = 44  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSpiral = 15  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSplit = 16  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectStretch = 17  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectStretchy = 45  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectStrips = 18  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectStyleEmphasis = 79  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSwish = 35  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectSwivel = 19  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectTeeter = 80  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectThinLine = 36  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectTransparency = 62  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectUnfold = 37  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectVerticalGrow = 81  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectWave = 82  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectWedge = 20  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectWheel = 21  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectWhip = 38  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectWipe = 22  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectZip = 46  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectZoom = 23  # Enum: MsoAnimEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffect
msoAnimEffectAfterFreeze = 1  # Enum: MsoAnimEffectAfter, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectafter
msoAnimEffectAfterHold = 3  # Enum: MsoAnimEffectAfter, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectafter
msoAnimEffectAfterRemove = 2  # Enum: MsoAnimEffectAfter, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectafter
msoAnimEffectAfterTransition = 4  # Enum: MsoAnimEffectAfter, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectafter
msoAnimEffectRestartAlways = 1  # Enum: MsoAnimEffectRestart, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectrestart
msoAnimEffectRestartNever = 3  # Enum: MsoAnimEffectRestart, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectrestart
msoAnimEffectRestartWhenOff = 2  # Enum: MsoAnimEffectRestart, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimeffectrestart
msoAnimFilterEffectSubtypeAcross = 9  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeDown = 25  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeDownLeft = 14  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeDownRight = 16  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeFromBottom = 13  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeFromLeft = 10  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeFromRight = 11  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeFromTop = 12  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeHorizontal = 5  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeIn = 7  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeInHorizontal = 3  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeInVertical = 1  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeLeft = 23  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeNone = 0  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeOut = 8  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeOutHorizontal = 4  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeOutVertical = 2  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeRight = 24  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeSpokes1 = 18  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeSpokes2 = 19  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeSpokes3 = 20  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeSpokes4 = 21  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeSpokes8 = 22  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeUp = 26  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeUpLeft = 15  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeUpRight = 17  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectSubtypeVertical = 6  # Enum: MsoAnimFilterEffectSubtype, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffectsubtype
msoAnimFilterEffectTypeBarn = 1  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeBlinds = 2  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeBox = 3  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeCheckerboard = 4  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeCircle = 5  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeDiamond = 6  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeDissolve = 7  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeFade = 8  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeImage = 9  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeNone = 0  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypePixelate = 10  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypePlus = 11  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeRandomBar = 12  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeSlide = 13  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeStretch = 14  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeStrips = 15  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeWedge = 16  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeWheel = 17  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimFilterEffectTypeWipe = 18  # Enum: MsoAnimFilterEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimfiltereffecttype
msoAnimColor = 7  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimHeight = 4  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimNone = 0  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimOpacity = 5  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimRotation = 6  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeFillBackColor = 1007  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeFillColor = 1005  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeFillOn = 1004  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeFillOpacity = 1006  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeLineColor = 1009  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeLineOn = 1008  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapePictureBrightness = 1001  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapePictureContrast = 1000  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapePictureGamma = 1002  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapePictureGrayscale = 1003  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeShadowColor = 1012  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeShadowOffsetX = 1014  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeShadowOffsetY = 1015  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeShadowOn = 1010  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeShadowOpacity = 1013  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimShapeShadowType = 1011  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletCharacter = 111  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletColor = 114  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletFontName = 112  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletNumber = 113  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletRelativeSize = 115  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletStyle = 116  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextBulletType = 117  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontBold = 100  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontColor = 101  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontEmboss = 102  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontItalic = 103  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontName = 104  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontShadow = 105  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontSize = 106  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontStrikeThrough = 110  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontSubscript = 107  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontSuperscript = 108  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextFontUnderline = 109  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimVisibility = 8  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimWidth = 3  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimX = 1  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimY = 2  # Enum: MsoAnimProperty, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimproperty
msoAnimTextUnitEffectByCharacter = 1  # Enum: MsoAnimTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtextuniteffect
msoAnimTextUnitEffectByParagraph = 0  # Enum: MsoAnimTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtextuniteffect
msoAnimTextUnitEffectByWord = 2  # Enum: MsoAnimTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtextuniteffect
msoAnimTextUnitEffectMixed = (
    -1
)  # Enum: MsoAnimTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtextuniteffect
msoAnimTriggerAfterPrevious = 3  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTriggerMixed = (
    -1
)  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTriggerNone = 0  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTriggerOnMediaBookmark = 5  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTriggerOnPageClick = 1  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTriggerOnShapeClick = 4  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTriggerWithPrevious = 2  # Enum: MsoAnimTriggerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtriggertype
msoAnimTypeColor = 2  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeCommand = 6  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeFilter = 7  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeMixed = (
    -2
)  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeMotion = 1  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeNone = 0  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeProperty = 5  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeRotation = 4  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeScale = 3  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimTypeSet = 8  # Enum: MsoAnimType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimtype
msoAnimateChartAllAtOnce = 7  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateChartByCategory = 8  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateChartByCategoryElements = 9  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateChartBySeries = 10  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateChartBySeriesElements = 11  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramAllAtOnce = 12  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramBreadthByLevel = 16  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramBreadthByNode = 15  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramClockwise = 17  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramClockwiseIn = 18  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramClockwiseOut = 19  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramCounterClockwise = 20  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramCounterClockwiseIn = 21  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramCounterClockwiseOut = 22  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramDepthByBranch = 14  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramDepthByNode = 13  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramDown = 26  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramInByRing = 23  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramOutByRing = 24  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateDiagramUp = 25  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateLevelMixed = (
    -1
)  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateLevelNone = 0  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateTextByAllLevels = 1  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateTextByFifthLevel = 6  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateTextByFirstLevel = 2  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateTextByFourthLevel = 5  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateTextBySecondLevel = 3  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoAnimateTextByThirdLevel = 4  # Enum: MsoAnimateByLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel
msoClickStateAfterAllAnimations = (
    -2
)  # Enum: MsoClickState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoclickstate
msoClickStateBeforeAutomaticAnimations = (
    -1
)  # Enum: MsoClickState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.msoclickstate
ppActionEndShow = 6  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionFirstSlide = 3  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionHyperlink = 7  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionLastSlide = 4  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionLastSlideViewed = 5  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionMixed = (
    -2
)  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionNamedSlideShow = 10  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionNextSlide = 1  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionNone = 0  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionOLEVerb = 11  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionPlay = 12  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionPreviousSlide = 2  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionRunMacro = 8  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppActionRunProgram = 9  # Enum: PpActionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppactiontype
ppAdvanceModeMixed = (
    -2
)  # Enum: PpAdvanceMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppadvancemode
ppAdvanceOnClick = 1  # Enum: PpAdvanceMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppadvancemode
ppAdvanceOnTime = 2  # Enum: PpAdvanceMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppadvancemode
ppAfterEffectDim = 2  # Enum: PpAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppaftereffect
ppAfterEffectHide = 1  # Enum: PpAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppaftereffect
ppAfterEffectHideOnClick = 3  # Enum: PpAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppaftereffect
ppAfterEffectMixed = (
    -2
)  # Enum: PpAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppaftereffect
ppAfterEffectNothing = 0  # Enum: PpAfterEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppaftereffect
ppAlertsAll = 2  # Enum: PpAlertLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppalertlevel
ppAlertsNone = 1  # Enum: PpAlertLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppalertlevel
ppArrangeCascade = 2  # Enum: PpArrangeStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pparrangestyle
ppArrangeTiled = 1  # Enum: PpArrangeStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pparrangestyle
ppAutoSizeMixed = (
    -2
)  # Enum: PpAutoSize, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppautosize
ppAutoSizeNone = 0  # Enum: PpAutoSize, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppautosize
ppAutoSizeShapeToFitText = 1  # Enum: PpAutoSize, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppautosize
ppBaselineAlignAuto = 5  # Enum: PpBaselineAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbaselinealignment
ppBaselineAlignBaseline = 1  # Enum: PpBaselineAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbaselinealignment
ppBaselineAlignCenter = 3  # Enum: PpBaselineAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbaselinealignment
ppBaselineAlignFarEast50 = 4  # Enum: PpBaselineAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbaselinealignment
ppBaselineAlignMixed = (
    -2
)  # Enum: PpBaselineAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbaselinealignment
ppBaselineAlignTop = 2  # Enum: PpBaselineAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbaselinealignment
ppBorderBottom = 3  # Enum: PpBorderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbordertype
ppBorderDiagonalDown = 5  # Enum: PpBorderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbordertype
ppBorderDiagonalUp = 6  # Enum: PpBorderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbordertype
ppBorderLeft = 2  # Enum: PpBorderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbordertype
ppBorderRight = 4  # Enum: PpBorderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbordertype
ppBorderTop = 1  # Enum: PpBorderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbordertype
ppBulletMixed = (
    -2
)  # Enum: PpBulletType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbullettype
ppBulletNone = 0  # Enum: PpBulletType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbullettype
ppBulletNumbered = 2  # Enum: PpBulletType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbullettype
ppBulletPicture = 3  # Enum: PpBulletType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbullettype
ppBulletUnnumbered = 1  # Enum: PpBulletType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppbullettype
ppCaseLower = 2  # Enum: PpChangeCase, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
ppCaseSentence = 1  # Enum: PpChangeCase, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
ppCaseTitle = 4  # Enum: PpChangeCase, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
ppCaseToggle = 5  # Enum: PpChangeCase, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
ppCaseUpper = 3  # Enum: PpChangeCase, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchangecase
ppAnimateByCategory = 2  # Enum: PpChartUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchartuniteffect
ppAnimateByCategoryElements = 4  # Enum: PpChartUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchartuniteffect
ppAnimateBySeries = 1  # Enum: PpChartUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchartuniteffect
ppAnimateBySeriesElements = 3  # Enum: PpChartUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchartuniteffect
ppAnimateChartAllAtOnce = 5  # Enum: PpChartUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchartuniteffect
ppAnimateChartMixed = (
    -2
)  # Enum: PpChartUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppchartuniteffect
ppCheckInMajorVersion = 1  # Enum: PpCheckInVersionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcheckinversiontype
ppCheckInMinorVersion = 0  # Enum: PpCheckInVersionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcheckinversiontype
ppCheckInOverwriteVersion = 2  # Enum: PpCheckInVersionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcheckinversiontype
ppAccent1 = 6  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppAccent2 = 7  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppAccent3 = 8  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppBackground = 1  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppFill = 5  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppForeground = 2  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppNotSchemeColor = 0  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppSchemeColorMixed = (
    -2
)  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppShadow = 3  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppTitle = 4  # Enum: PpColorSchemeIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppcolorschemeindex
ppDateTimeFigureOut = 14  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeFormatMixed = (
    -2
)  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeHmm = 10  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeHmmss = 11  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeMMMMdyyyy = 4  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeMMMMyy = 6  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeMMddyyHmm = 8  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeMMddyyhmmAMPM = 9  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeMMyy = 7  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeMdyy = 1  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ1 = 15  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ2 = 16  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ3 = 17  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ4 = 18  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ5 = 19  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ6 = 20  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeUAQ7 = 21  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimedMMMMyyyy = 3  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimedMMMyy = 5  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimeddddMMMMddyyyy = 2  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimehmmAMPM = 12  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDateTimehmmssAMPM = 13  # Enum: PpDateTimeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdatetimeformat
ppDirectionLeftToRight = 1  # Enum: PpDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdirection
ppDirectionMixed = (
    -2
)  # Enum: PpDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdirection
ppDirectionRightToLeft = 2  # Enum: PpDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppdirection
ppEffectAirplaneLeft = 3950  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectAirplaneRight = 3951  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectAppear = 3844  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBlindsHorizontal = 769  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBlindsVertical = 770  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBoxDown = 3925  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBoxIn = 3074  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBoxLeft = 3922  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBoxOut = 3073  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBoxRight = 3924  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectBoxUp = 3923  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCheckerboardAcross = 1025  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCheckerboardDown = 1026  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCircleOut = 3845  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCombHorizontal = 3847  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCombVertical = 3848  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectConveyorLeft = 3882  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectConveyorRight = 3883  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverDown = 1284  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverLeft = 1281  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverLeftDown = 1287  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverLeftUp = 1285  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverRight = 1283  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverRightDown = 1288  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverRightUp = 1286  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCoverUp = 1282  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCrawlFromDown = 3344  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCrawlFromLeft = 3341  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCrawlFromRight = 3343  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCrawlFromUp = 3342  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCrush = 3943  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCubeDown = 3917  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCubeLeft = 3914  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCubeRight = 3916  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCubeUp = 3915  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCurtains = 3938  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCut = 257  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectCutThroughBlack = 258  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectDiamondOut = 3846  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectDissolve = 1537  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectDoorsHorizontal = 3885  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectDoorsVertical = 3884  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectDrapeLeft = 3936  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectDrapeRight = 3937  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFade = 1793  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFadeSmoothly = 3849  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFallOverLeft = 3934  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFallOverRight = 3935  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFerrisWheelLeft = 3899  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFerrisWheelRight = 3900  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlashOnceFast = 3841  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlashOnceMedium = 3842  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlashOnceSlow = 3843  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlashbulb = 3909  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlipDown = 3908  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlipLeft = 3905  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlipRight = 3907  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlipUp = 3906  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromBottom = 3332  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromBottomLeft = 3335  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromBottomRight = 3336  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromLeft = 3329  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromRight = 3331  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromTop = 3330  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromTopLeft = 3333  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyFromTopRight = 3334  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyThroughIn = 3890  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyThroughInBounce = 3892  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyThroughOut = 3891  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFlyThroughOutBounce = 3893  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectFracture = 3942  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGalleryLeft = 3880  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGalleryRight = 3881  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterDiamondDown = 3875  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterDiamondLeft = 3872  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterDiamondRight = 3874  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterDiamondUp = 3873  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterHexagonDown = 3879  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterHexagonLeft = 3876  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterHexagonRight = 3878  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectGlitterHexagonUp = 3877  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectHoneycomb = 3898  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectMixed = (
    -2
)  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectMorphByChar = 3956  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectMorphByObject = 3954  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectMorphByWord = 3955  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectNewsflash = 3850  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectNone = 0  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectOrbitDown = 3929  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectOrbitLeft = 3926  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectOrbitRight = 3928  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectOrbitUp = 3927  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectOrigamiLeft = 3952  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectOrigamiRight = 3953  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPageCurlDoubleLeft = 3948  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPageCurlDoubleRight = 3949  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPageCurlSingleLeft = 3946  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPageCurlSingleRight = 3947  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPanDown = 3933  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPanLeft = 3930  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPanRight = 3932  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPanUp = 3931  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPeekFromDown = 3338  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPeekFromLeft = 3337  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPeekFromRight = 3339  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPeekFromUp = 3340  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPeelOffLeft = 3944  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPeelOffRight = 3945  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPlusOut = 3851  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPrestige = 3941  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPushDown = 3852  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPushLeft = 3853  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPushRight = 3854  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectPushUp = 3855  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRandom = 513  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRandomBarsHorizontal = 2305  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRandomBarsVertical = 2306  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRevealBlackLeft = 3896  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRevealBlackRight = 3897  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRevealSmoothLeft = 3894  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRevealSmoothRight = 3895  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRippleCenter = 3867  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRippleLeftDown = 3870  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRippleLeftUp = 3869  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRippleRightDown = 3871  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRippleRightUp = 3868  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRotateDown = 3921  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRotateLeft = 3918  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRotateRight = 3920  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectRotateUp = 3919  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectShredRectangleIn = 3912  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectShredRectangleOut = 3913  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectShredStripsIn = 3910  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectShredStripsOut = 3911  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSpiral = 3357  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSplitHorizontalIn = 3586  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSplitHorizontalOut = 3585  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSplitVerticalIn = 3588  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSplitVerticalOut = 3587  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStretchAcross = 3351  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStretchDown = 3355  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStretchLeft = 3352  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStretchRight = 3354  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStretchUp = 3353  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsDownLeft = 2563  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsDownRight = 2564  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsLeftDown = 2567  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsLeftUp = 2565  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsRightDown = 2568  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsRightUp = 2566  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsUpLeft = 2561  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectStripsUpRight = 2562  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSwitchDown = 3904  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSwitchLeft = 3901  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSwitchRight = 3903  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSwitchUp = 3902  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectSwivel = 3356  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverDown = 2052  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverLeft = 2049  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverLeftDown = 2055  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverLeftUp = 2053  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverRight = 2051  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverRightDown = 2056  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverRightUp = 2054  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectUncoverUp = 2050  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectVortexDown = 3866  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectVortexLeft = 3863  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectVortexRight = 3865  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectVortexUp = 3864  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWarpIn = 3888  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWarpOut = 3889  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWedge = 3856  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWheel1Spoke = 3857  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWheel2Spokes = 3858  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWheel3Spokes = 3859  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWheel4Spokes = 3860  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWheel8Spokes = 3861  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWheelReverse1Spoke = 3862  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWindLeft = 3939  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWindRight = 3940  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWindowHorizontal = 3887  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWindowVertical = 3886  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWipeDown = 2820  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWipeLeft = 2817  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWipeRight = 2819  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectWipeUp = 2818  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectZoomBottom = 3350  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectZoomCenter = 3349  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectZoomIn = 3345  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectZoomInSlightly = 3346  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectZoomOut = 3347  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppEffectZoomOutSlightly = 3348  # Enum: PpEntryEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppentryeffect
ppClipRelativeToSlide = 2  # Enum: PpExportMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppexportmode
ppRelativeToSlide = 1  # Enum: PpExportMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppexportmode
ppScaleToFit = 3  # Enum: PpExportMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppexportmode
ppScaleXY = 4  # Enum: PpExportMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppexportmode
ppFarEastLineBreakLevelCustom = 3  # Enum: PpFarEastLineBreakLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfareastlinebreaklevel
ppFarEastLineBreakLevelNormal = 1  # Enum: PpFarEastLineBreakLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfareastlinebreaklevel
ppFarEastLineBreakLevelStrict = 2  # Enum: PpFarEastLineBreakLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfareastlinebreaklevel
ppFileDialogOpen = 1  # Enum: PpFileDialogType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfiledialogtype
ppFileDialogSave = 2  # Enum: PpFileDialogType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfiledialogtype
ppFixedFormatIntentPrint = 2  # Enum: PpFixedFormatIntent, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfixedformatintent
ppFixedFormatIntentScreen = 1  # Enum: PpFixedFormatIntent, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfixedformatintent
ppFixedFormatTypePDF = 2  # Enum: PpFixedFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfixedformattype
ppFixedFormatTypeXPS = 1  # Enum: PpFixedFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfixedformattype
ppFollowColorsMixed = (
    -2
)  # Enum: PpFollowColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfollowcolors
ppFollowColorsNone = 0  # Enum: PpFollowColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfollowcolors
ppFollowColorsScheme = 1  # Enum: PpFollowColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfollowcolors
ppFollowColorsTextAndBackground = 2  # Enum: PpFollowColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppfollowcolors
ppFrameColorsBlackTextOnWhite = 5  # Enum: PpFrameColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppframecolors
ppFrameColorsBrowserColors = 1  # Enum: PpFrameColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppframecolors
ppFrameColorsPresentationSchemeAccentColor = 3  # Enum: PpFrameColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppframecolors
ppFrameColorsPresentationSchemeTextColor = 2  # Enum: PpFrameColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppframecolors
ppFrameColorsWhiteTextOnBlack = 4  # Enum: PpFrameColors, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppframecolors
ppHorizontalGuide = 1  # Enum: PpGuideOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppguideorientation
ppVerticalGuide = 2  # Enum: PpGuideOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppguideorientation
ppHTMLAutodetect = 4  # Enum: PpHTMLVersion, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pphtmlversion
ppHTMLDual = 3  # Enum: PpHTMLVersion, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pphtmlversion
ppHTMLv3 = 1  # Enum: PpHTMLVersion, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pphtmlversion
ppHTMLv4 = 2  # Enum: PpHTMLVersion, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pphtmlversion
ppIndentControlMixed = (
    -2
)  # Enum: PpIndentControl, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppindentcontrol
ppIndentKeepAttr = 2  # Enum: PpIndentControl, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppindentcontrol
ppIndentReplaceAttr = 1  # Enum: PpIndentControl, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppindentcontrol
ppMediaTaskStatusDone = 3  # Enum: PpMediaTaskStatus, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediataskstatus
ppMediaTaskStatusFailed = 4  # Enum: PpMediaTaskStatus, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediataskstatus
ppMediaTaskStatusInProgress = 1  # Enum: PpMediaTaskStatus, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediataskstatus
ppMediaTaskStatusNone = 0  # Enum: PpMediaTaskStatus, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediataskstatus
ppMediaTaskStatusQueued = 2  # Enum: PpMediaTaskStatus, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediataskstatus
ppMediaTypeMixed = (
    -2
)  # Enum: PpMediaType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediatype
ppMediaTypeMovie = 3  # Enum: PpMediaType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediatype
ppMediaTypeOther = 1  # Enum: PpMediaType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediatype
ppMediaTypeSound = 2  # Enum: PpMediaType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmediatype
ppMouseClick = 1  # Enum: PpMouseActivation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmouseactivation
ppMouseOver = 2  # Enum: PpMouseActivation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppmouseactivation
ppBulletAlphaLCParenBoth = 8  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletAlphaLCParenRight = 9  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletAlphaLCPeriod = 0  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletAlphaUCParenBoth = 10  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletAlphaUCParenRight = 11  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletAlphaUCPeriod = 1  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicAbjadDash = 24  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicAlphaDash = 23  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicDBPeriod = 29  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicDBPlain = 28  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicParenBoth = 12  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicParenRight = 2  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicPeriod = 3  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletArabicPlain = 13  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletCircleNumDBPlain = 18  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletCircleNumWDBlackPlain = 20  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletCircleNumWDWhitePlain = 19  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletHebrewAlphaDash = 25  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletHindiAlpha1Period = 40  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletHindiAlphaPeriod = 36  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletHindiNumParenRight = 39  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletHindiNumPeriod = 37  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletKanjiKoreanPeriod = 27  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletKanjiKoreanPlain = 26  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletKanjiSimpChinDBPeriod = 38  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletRomanLCParenBoth = 4  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletRomanLCParenRight = 5  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletRomanLCPeriod = 6  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletRomanUCParenBoth = 14  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletRomanUCParenRight = 15  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletRomanUCPeriod = 7  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletSimpChinPeriod = 17  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletSimpChinPlain = 16  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletStyleMixed = (
    -2
)  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletThaiAlphaParenBoth = 32  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletThaiAlphaParenRight = 31  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletThaiAlphaPeriod = 30  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletThaiNumParenBoth = 35  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletThaiNumParenRight = 34  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletThaiNumPeriod = 33  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletTradChinPeriod = 22  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppBulletTradChinPlain = 21  # Enum: PpNumberedBulletStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppnumberedbulletstyle
ppAlignCenter = 2  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignDistribute = 5  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignJustify = 4  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignJustifyLow = 7  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignLeft = 1  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignRight = 3  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignThaiDistribute = 6  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppAlignmentMixed = (
    -2
)  # Enum: PpParagraphAlignment, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppparagraphalignment
ppPasteBitmap = 1  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteDefault = 0  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteEnhancedMetafile = 2  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteGIF = 4  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteHTML = 8  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteJPG = 5  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteMetafilePicture = 3  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteOLEObject = 10  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPastePNG = 6  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteRTF = 9  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteSVG = 12  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteShape = 11  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPasteText = 7  # Enum: PpPasteDataType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppastedatatype
ppPlaceholderBitmap = 9  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderBody = 2  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderCenterTitle = 3  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderChart = 8  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderDate = 16  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderFooter = 15  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderHeader = 14  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderMediaClip = 10  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderMixed = (
    -2
)  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderObject = 7  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderOrgChart = 11  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderPicture = 18  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderSlideNumber = 13  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderSubtitle = 4  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderTable = 12  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderTitle = 1  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderVerticalBody = 6  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderVerticalObject = 17  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderVerticalTitle = 5  # Enum: PpPlaceholderType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppNotReady = 3  # Enum: PpPlayerState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplayerstate
ppPaused = 1  # Enum: PpPlayerState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplayerstate
ppPlaying = 0  # Enum: PpPlayerState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplayerstate
ppStopped = 2  # Enum: PpPlayerState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplayerstate
ppPrintBlackAndWhite = 2  # Enum: PpPrintColorType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintcolortype
ppPrintColor = 1  # Enum: PpPrintColorType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintcolortype
ppPrintPureBlackAndWhite = 3  # Enum: PpPrintColorType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintcolortype
ppPrintHandoutHorizontalFirst = 2  # Enum: PpPrintHandoutOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprinthandoutorder
ppPrintHandoutVerticalFirst = 1  # Enum: PpPrintHandoutOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprinthandoutorder
ppPrintOutputBuildSlides = 7  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputFourSlideHandouts = 8  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputNineSlideHandouts = 9  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputNotesPages = 5  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputOneSlideHandouts = 10  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputOutline = 6  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputSixSlideHandouts = 4  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputSlides = 1  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputThreeSlideHandouts = 3  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintOutputTwoSlideHandouts = 2  # Enum: PpPrintOutputType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintoutputtype
ppPrintAll = 1  # Enum: PpPrintRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
ppPrintCurrent = 3  # Enum: PpPrintRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
ppPrintNamedSlideShow = 5  # Enum: PpPrintRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
ppPrintSection = 6  # Enum: PpPrintRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
ppPrintSelection = 2  # Enum: PpPrintRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
ppPrintSlideRange = 4  # Enum: PpPrintRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrangetype
ppProtectedViewCloseEdit = 1  # Enum: PpProtectedViewCloseReason, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprotectedviewclosereason
ppProtectedViewCloseForced = 2  # Enum: PpProtectedViewCloseReason, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprotectedviewclosereason
ppProtectedViewCloseNormal = 0  # Enum: PpProtectedViewCloseReason, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppprotectedviewclosereason
ppPublishAll = 1  # Enum: PpPublishSourceType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppublishsourcetype
ppPublishNamedSlideShow = 3  # Enum: PpPublishSourceType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppublishsourcetype
ppPublishSlideRange = 2  # Enum: PpPublishSourceType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pppublishsourcetype
ppRDIAll = 99  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIAtMentions = 18  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIComments = 1  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIContentType = 16  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIDocumentManagementPolicy = 15  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIDocumentProperties = 8  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIDocumentServerProperties = 14  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIDocumentWorkspace = 10  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIInkAnnotations = 11  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIPublishPath = 13  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDIRemovePersonalInformation = 4  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppRDISlideUpdateInformation = 17  # Enum: PpRemoveDocInfoType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppremovedocinfotype
ppResampleMediaProfileCustom = 1  # Enum: PpResampleMediaProfile, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppresamplemediaprofile
ppResampleMediaProfileSmall = 2  # Enum: PpResampleMediaProfile, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppresamplemediaprofile
ppResampleMediaProfileSmaller = 3  # Enum: PpResampleMediaProfile, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppresamplemediaprofile
ppResampleMediaProfileSmallest = 4  # Enum: PpResampleMediaProfile, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppresamplemediaprofile
ppRevisionInfoBaseline = 1  # Enum: PpRevisionInfo, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pprevisioninfo
ppRevisionInfoMerged = 2  # Enum: PpRevisionInfo, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pprevisioninfo
ppRevisionInfoNone = 0  # Enum: PpRevisionInfo, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pprevisioninfo
ppSaveAsAddIn = 8  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsAnimatedGIF = 40  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsBMP = 19  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsDefault = 11  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsEMF = 23  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsExternalConverter = 64000  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsGIF = 16  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsHTML = 12  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsHTMLDual = 14  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsHTMLv3 = 13  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsJPG = 17  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsMP4 = 39  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsMetaFile = 15  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenDocumentPresentation = 35  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLAddin = 30  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLPicturePresentation = 36  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLPresentation = 24  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLPresentationMacroEnabled = 25  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLShow = 28  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLShowMacroEnabled = 29  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLTemplate = 26  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLTemplateMacroEnabled = 27  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsOpenXMLTheme = 31  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPDF = 32  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPNG = 18  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPowerPoint3 = 4  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPowerPoint4 = 3  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPowerPoint4FarEast = 10  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPowerPoint7 = 2  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPresForReview = 22  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsPresentation = 1  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsRTF = 6  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsShow = 7  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsStrictOpenXMLPresentation = 38  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsTIF = 21  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsTemplate = 5  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsWMV = 37  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsWebArchive = 20  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsXMLPresentation = 34  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSaveAsXPS = 33  # Enum: PpSaveAsFileType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
ppSelectionNone = 0  # Enum: PpSelectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppselectiontype
ppSelectionShapes = 2  # Enum: PpSelectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppselectiontype
ppSelectionSlides = 1  # Enum: PpSelectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppselectiontype
ppSelectionText = 3  # Enum: PpSelectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppselectiontype
ppShapeFormatBMP = 3  # Enum: PpShapeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppshapeformat
ppShapeFormatEMF = 5  # Enum: PpShapeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppshapeformat
ppShapeFormatGIF = 0  # Enum: PpShapeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppshapeformat
ppShapeFormatJPG = 1  # Enum: PpShapeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppshapeformat
ppShapeFormatPNG = 2  # Enum: PpShapeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppshapeformat
ppShapeFormatWMF = 4  # Enum: PpShapeFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppshapeformat
ppLayoutBlank = 12  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutChart = 8  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutChartAndText = 6  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutClipArtAndVerticalText = 26  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutClipartAndText = 10  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutComparison = 34  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutContentWithCaption = 35  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutCustom = 32  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutFourObjects = 24  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutLargeObject = 15  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutMediaClipAndText = 18  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutMixed = (
    -2
)  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutObject = 16  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutObjectAndText = 14  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutObjectAndTwoObjects = 30  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutObjectOverText = 19  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutOrgchart = 7  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutPictureWithCaption = 36  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutSectionHeader = 33  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTable = 4  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutText = 2  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTextAndChart = 5  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTextAndClipart = 9  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTextAndMediaClip = 17  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTextAndObject = 13  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTextAndTwoObjects = 21  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTextOverObject = 20  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTitle = 1  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTitleOnly = 11  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTwoColumnText = 3  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTwoObjects = 29  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTwoObjectsAndObject = 31  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTwoObjectsAndText = 22  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutTwoObjectsOverText = 23  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutVerticalText = 25  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutVerticalTitleAndText = 27  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutVerticalTitleAndTextOverChart = 28  # Enum: PpSlideLayout, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppSlideShowManualAdvance = 1  # Enum: PpSlideShowAdvanceMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowadvancemode
ppSlideShowRehearseNewTimings = 3  # Enum: PpSlideShowAdvanceMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowadvancemode
ppSlideShowUseSlideTimings = 2  # Enum: PpSlideShowAdvanceMode, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowadvancemode
ppSlideShowPointerAlwaysHidden = 3  # Enum: PpSlideShowPointerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowpointertype
ppSlideShowPointerArrow = 1  # Enum: PpSlideShowPointerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowpointertype
ppSlideShowPointerAutoArrow = 4  # Enum: PpSlideShowPointerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowpointertype
ppSlideShowPointerEraser = 5  # Enum: PpSlideShowPointerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowpointertype
ppSlideShowPointerNone = 0  # Enum: PpSlideShowPointerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowpointertype
ppSlideShowPointerPen = 2  # Enum: PpSlideShowPointerType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowpointertype
ppShowAll = 1  # Enum: PpSlideShowRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowrangetype
ppShowNamedSlideShow = 3  # Enum: PpSlideShowRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowrangetype
ppShowSlideRange = 2  # Enum: PpSlideShowRangeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowrangetype
ppSlideShowBlackScreen = 3  # Enum: PpSlideShowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowstate
ppSlideShowDone = 5  # Enum: PpSlideShowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowstate
ppSlideShowPaused = 2  # Enum: PpSlideShowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowstate
ppSlideShowRunning = 1  # Enum: PpSlideShowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowstate
ppSlideShowWhiteScreen = 4  # Enum: PpSlideShowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowstate
ppShowTypeKiosk = 3  # Enum: PpSlideShowType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowtype
ppShowTypeSpeaker = 1  # Enum: PpSlideShowType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowtype
ppShowTypeWindow = 2  # Enum: PpSlideShowType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowtype
ppShowTypeWindow2 = 4  # Enum: PpSlideShowType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslideshowtype
ppSlideSize35MM = 4  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeA3Paper = 9  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeA4Paper = 3  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeB4ISOPaper = 10  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeB4JISPaper = 12  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeB5ISOPaper = 11  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeB5JISPaper = 13  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeBanner = 6  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeCustom = 7  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeHagakiCard = 14  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeLedgerPaper = 8  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeLetterPaper = 2  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeOnScreen = 1  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeOnScreen16x10 = 16  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeOnScreen16x9 = 15  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSlideSizeOverhead = 5  # Enum: PpSlideSizeType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidesizetype
ppSoundEffectsMixed = (
    -2
)  # Enum: PpSoundEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundeffecttype
ppSoundFile = 2  # Enum: PpSoundEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundeffecttype
ppSoundNone = 0  # Enum: PpSoundEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundeffecttype
ppSoundStopPrevious = 1  # Enum: PpSoundEffectType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundeffecttype
ppSoundFormatCDAudio = 3  # Enum: PpSoundFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundformattype
ppSoundFormatMIDI = 2  # Enum: PpSoundFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundformattype
ppSoundFormatMixed = (
    -2
)  # Enum: PpSoundFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundformattype
ppSoundFormatNone = 0  # Enum: PpSoundFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundformattype
ppSoundFormatWAV = 1  # Enum: PpSoundFormatType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsoundformattype
ppTabStopCenter = 2  # Enum: PpTabStopType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptabstoptype
ppTabStopDecimal = 4  # Enum: PpTabStopType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptabstoptype
ppTabStopLeft = 1  # Enum: PpTabStopType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptabstoptype
ppTabStopMixed = (
    -2
)  # Enum: PpTabStopType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptabstoptype
ppTabStopRight = 3  # Enum: PpTabStopType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptabstoptype
ppAnimateByAllLevels = 16  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateByFifthLevel = 5  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateByFirstLevel = 1  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateByFourthLevel = 4  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateBySecondLevel = 2  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateByThirdLevel = 3  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateLevelMixed = (
    -2
)  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppAnimateLevelNone = 0  # Enum: PpTextLevelEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextleveleffect
ppBodyStyle = 3  # Enum: PpTextStyleType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextstyletype
ppDefaultStyle = 1  # Enum: PpTextStyleType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextstyletype
ppTitleStyle = 2  # Enum: PpTextStyleType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextstyletype
ppAnimateByCharacter = 2  # Enum: PpTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextuniteffect
ppAnimateByParagraph = 0  # Enum: PpTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextuniteffect
ppAnimateByWord = 1  # Enum: PpTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextuniteffect
ppAnimateUnitMixed = (
    -2
)  # Enum: PpTextUnitEffect, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptextuniteffect
ppTransitionSpeedFast = 3  # Enum: PpTransitionSpeed, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptransitionspeed
ppTransitionSpeedMedium = 2  # Enum: PpTransitionSpeed, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptransitionspeed
ppTransitionSpeedMixed = (
    -2
)  # Enum: PpTransitionSpeed, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptransitionspeed
ppTransitionSpeedSlow = 1  # Enum: PpTransitionSpeed, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.pptransitionspeed
ppUpdateOptionAutomatic = 2  # Enum: PpUpdateOption, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppupdateoption
ppUpdateOptionManual = 1  # Enum: PpUpdateOption, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppupdateoption
ppUpdateOptionMixed = (
    -2
)  # Enum: PpUpdateOption, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppupdateoption
ppViewHandoutMaster = 4  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewMasterThumbnails = 12  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewNormal = 9  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewNotesMaster = 5  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewNotesPage = 3  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewOutline = 6  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewPrintPreview = 10  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewSlide = 1  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewSlideMaster = 2  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewSlideSorter = 7  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewThumbnails = 11  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppViewTitleMaster = 8  # Enum: PpViewType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppviewtype
ppWindowMaximized = 3  # Enum: PpWindowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppwindowstate
ppWindowMinimized = 2  # Enum: PpWindowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppwindowstate
ppWindowNormal = 1  # Enum: PpWindowState, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppwindowstate
xlAxisCrossesAutomatic = (
    -4105
)  # Enum: XlAxisCrosses, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxiscrosses
xlAxisCrossesCustom = (
    -4114
)  # Enum: XlAxisCrosses, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxiscrosses
xlAxisCrossesMaximum = 2  # Enum: XlAxisCrosses, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxiscrosses
xlAxisCrossesMinimum = 4  # Enum: XlAxisCrosses, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxiscrosses
xlPrimary = 1  # Enum: XlAxisGroup, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxisgroup
xlSecondary = 2  # Enum: XlAxisGroup, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxisgroup
xlCategory = 1  # Enum: XlAxisType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxistype
xlSeriesAxis = 3  # Enum: XlAxisType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxistype
xlValue = 2  # Enum: XlAxisType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlaxistype
xlBackgroundAutomatic = (
    -4105
)  # Enum: XlBackground, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbackground
xlBackgroundOpaque = 3  # Enum: XlBackground, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbackground
xlBackgroundTransparent = 2  # Enum: XlBackground, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbackground
xlBox = 0  # Enum: XlBarShape, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbarshape
xlConeToMax = 5  # Enum: XlBarShape, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbarshape
xlConeToPoint = 4  # Enum: XlBarShape, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbarshape
xlCylinder = 3  # Enum: XlBarShape, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbarshape
xlPyramidToMax = 2  # Enum: XlBarShape, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbarshape
xlPyramidToPoint = 1  # Enum: XlBarShape, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbarshape
xlBinsTypeAutomatic = 0  # Enum: XlBinsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbinstype
xlBinsTypeBinCount = 4  # Enum: XlBinsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbinstype
xlBinsTypeBinSize = 3  # Enum: XlBinsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbinstype
xlBinsTypeCategorical = 1  # Enum: XlBinsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbinstype
xlBinsTypeManual = 2  # Enum: XlBinsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlbinstype
xlHairline = 1  # Enum: XlBorderWeight, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlborderweight
xlMedium = (
    -4138
)  # Enum: XlBorderWeight, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlborderweight
xlThick = 4  # Enum: XlBorderWeight, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlborderweight
xlThin = 2  # Enum: XlBorderWeight, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlborderweight
xlCategoryLabelLevelAll = (
    -1
)  # Enum: XlCategoryLabelLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorylabellevel
xlCategoryLabelLevelCustom = (
    -2
)  # Enum: XlCategoryLabelLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorylabellevel
xlCategoryLabelLevelNone = (
    -3
)  # Enum: XlCategoryLabelLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorylabellevel
xlCategoryAscending = 2  # Enum: XlCategorySortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorysortorder
xlCategoryDescending = 3  # Enum: XlCategorySortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorysortorder
xlIndexAscending = 0  # Enum: XlCategorySortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorysortorder
xlIndexDescending = 1  # Enum: XlCategorySortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorysortorder
xlAutomaticScale = (
    -4105
)  # Enum: XlCategoryType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorytype
xlCategoryScale = 2  # Enum: XlCategoryType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorytype
xlTimeScale = 3  # Enum: XlCategoryType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcategorytype
xlChartElementPositionAutomatic = (
    -4105
)  # Enum: XlChartElementPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartelementposition
xlChartElementPositionCustom = (
    -4114
)  # Enum: XlChartElementPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartelementposition
xlAnyGallery = 23  # Enum: XlChartGallery, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartgallery
xlBuiltIn = 21  # Enum: XlChartGallery, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartgallery
xlUserDefined = 22  # Enum: XlChartGallery, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartgallery
xlAxis = 21  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlAxisTitle = 17  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlChartArea = 2  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlChartTitle = 4  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlCorners = 6  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlDataLabel = 0  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlDataTable = 7  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlDisplayUnitLabel = 30  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlDownBars = 20  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlDropLines = 26  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlErrorBars = 9  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlFloor = 23  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlHiLoLines = 25  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlLeaderLines = 29  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlLegend = 24  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlLegendEntry = 12  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlLegendKey = 13  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlMajorGridlines = 15  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlMinorGridlines = 16  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlNothing = 28  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlPivotChartDropZone = 32  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlPivotChartFieldButton = 31  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlPlotArea = 19  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlRadarAxisLabels = 27  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlSeries = 3  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlSeriesLines = 22  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlShape = 14  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlTrendline = 8  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlUpBars = 18  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlWalls = 5  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlXErrorBars = 10  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlYErrorBars = 11  # Enum: XlChartItem, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartitem
xlAllFaces = 7  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlEnd = 2  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlEndSides = 3  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlFront = 4  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlFrontEnd = 6  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlFrontSides = 5  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlSides = 1  # Enum: XlChartPicturePlacement, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpictureplacement
xlStack = 2  # Enum: XlChartPictureType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpicturetype
xlStackScale = 3  # Enum: XlChartPictureType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpicturetype
xlStretch = 1  # Enum: XlChartPictureType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartpicturetype
xlSplitByCustomSplit = 4  # Enum: XlChartSplitType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartsplittype
xlSplitByPercentValue = 3  # Enum: XlChartSplitType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartsplittype
xlSplitByPosition = 1  # Enum: XlChartSplitType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartsplittype
xlSplitByValue = 2  # Enum: XlChartSplitType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlchartsplittype
xlColorIndexAutomatic = (
    -4105
)  # Enum: XlColorIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcolorindex
xlColorIndexNone = (
    -4142
)  # Enum: XlColorIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcolorindex
xl3DBar = (
    -4099
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xl3DSurface = (
    -4103
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlAbove = 0  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlAutomatic = (
    -4105
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlBar = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlBelow = 1  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlBoth = 1  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlBottom = (
    -4107
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCenter = (
    -4108
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlChecker = 9  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCircle = 8  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlColumn = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCombination = (
    -4111
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCorner = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCrissCross = 16  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCross = 4  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlCustom = (
    -4114
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlDefaultAutoFormat = (
    -1
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlDiamond = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlDistributed = (
    -4117
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlFill = 5  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlFixedValue = 1  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGeneral = 1  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGray16 = 17  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGray25 = (
    -4124
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGray50 = (
    -4125
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGray75 = (
    -4126
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGray8 = 18  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlGrid = 15  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlHigh = (
    -4127
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlInside = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlJustify = (
    -4130
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlLeft = (
    -4131
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlLightDown = 13  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlLightHorizontal = 11  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlLightUp = 14  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlLightVertical = 12  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlLow = (
    -4134
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlMaximum = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlMinimum = 4  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlMinusValues = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlNextToAxis = 4  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlNone = (
    -4142
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlOpaque = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlOutside = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlPercent = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlPlus = 9  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlPlusValues = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlRight = (
    -4152
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlScale = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlSemiGray75 = 10  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlShowLabel = 4  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlShowLabelAndPercent = 5  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlShowPercent = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlShowValue = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlSingle = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlSolid = 1  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlSquare = 1  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlStError = 4  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlStar = 5  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlTop = (
    -4160
)  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlTransparent = 2  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlTriangle = 3  # Enum: XlConstants, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlconstants
xlBitmap = 2  # Enum: XlCopyPictureFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcopypictureformat
xlPicture = (
    -4147
)  # Enum: XlCopyPictureFormat, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlcopypictureformat
xlLabelPositionAbove = 0  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionBelow = 1  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionBestFit = 5  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionCenter = (
    -4108
)  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionCustom = 7  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionInsideBase = 4  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionInsideEnd = 3  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionLeft = (
    -4131
)  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionMixed = 6  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionOutsideEnd = 2  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlLabelPositionRight = (
    -4152
)  # Enum: XlDataLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelposition
xlDataLabelSeparatorDefault = 1  # Enum: XlDataLabelSeparator, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelseparator
xlDataLabelsShowBubbleSizes = 6  # Enum: XlDataLabelsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelstype
xlDataLabelsShowLabel = 4  # Enum: XlDataLabelsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelstype
xlDataLabelsShowLabelAndPercent = 5  # Enum: XlDataLabelsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelstype
xlDataLabelsShowNone = (
    -4142
)  # Enum: XlDataLabelsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelstype
xlDataLabelsShowPercent = 3  # Enum: XlDataLabelsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelstype
xlDataLabelsShowValue = 2  # Enum: XlDataLabelsType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldatalabelstype
xlInterpolated = 3  # Enum: XlDisplayBlanksAs, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayblanksas
xlNotPlotted = 1  # Enum: XlDisplayBlanksAs, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayblanksas
xlZero = 2  # Enum: XlDisplayBlanksAs, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayblanksas
xlHundredMillions = (
    -8
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlHundredThousands = (
    -5
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlHundreds = (
    -2
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlMillionMillions = (
    -10
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlMillions = (
    -6
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlTenMillions = (
    -7
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlTenThousands = (
    -4
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlThousandMillions = (
    -9
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlThousands = (
    -3
)  # Enum: XlDisplayUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xldisplayunit
xlCap = 1  # Enum: XlEndStyleCap, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlendstylecap
xlNoCap = 2  # Enum: XlEndStyleCap, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlendstylecap
xlChartX = (
    -4168
)  # Enum: XlErrorBarDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbardirection
xlChartY = 1  # Enum: XlErrorBarDirection, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbardirection
xlErrorBarIncludeBoth = 1  # Enum: XlErrorBarInclude, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbarinclude
xlErrorBarIncludeMinusValues = 3  # Enum: XlErrorBarInclude, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbarinclude
xlErrorBarIncludeNone = (
    -4142
)  # Enum: XlErrorBarInclude, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbarinclude
xlErrorBarIncludePlusValues = 2  # Enum: XlErrorBarInclude, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbarinclude
xlErrorBarTypeCustom = (
    -4114
)  # Enum: XlErrorBarType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbartype
xlErrorBarTypeFixedValue = 1  # Enum: XlErrorBarType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbartype
xlErrorBarTypePercent = 2  # Enum: XlErrorBarType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbartype
xlErrorBarTypeStDev = (
    -4155
)  # Enum: XlErrorBarType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbartype
xlErrorBarTypeStError = 4  # Enum: XlErrorBarType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlerrorbartype
xlGeoMappingLevelAutomatic = 0  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelCountryRegion = 5  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelCountryRegionList = 6  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelCounty = 3  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelDataOnly = 1  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelPostalCode = 2  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelState = 4  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoMappingLevelWorld = 7  # Enum: XlGeoMappingLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeomappinglevel
xlGeoProjectionTypeAlbers = 3  # Enum: XlGeoProjectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeoprojectiontype
xlGeoProjectionTypeAutomatic = 0  # Enum: XlGeoProjectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeoprojectiontype
xlGeoProjectionTypeMercator = 1  # Enum: XlGeoProjectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeoprojectiontype
xlGeoProjectionTypeMiller = 2  # Enum: XlGeoProjectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeoprojectiontype
xlGeoProjectionTypeRobinson = 4  # Enum: XlGeoProjectionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgeoprojectiontype
xlGradientStopPositionTypeExtremeValue = 0  # Enum: XlGradientStopPositionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgradientstoppositiontype
xlGradientStopPositionTypeNumber = 1  # Enum: XlGradientStopPositionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgradientstoppositiontype
xlGradientStopPositionTypePercent = 2  # Enum: XlGradientStopPositionType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlgradientstoppositiontype
xlHAlignCenter = (
    -4108
)  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignCenterAcrossSelection = 7  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignDistributed = (
    -4117
)  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignFill = 5  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignGeneral = 1  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignJustify = (
    -4130
)  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignLeft = (
    -4131
)  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlHAlignRight = (
    -4152
)  # Enum: XlHAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlhalign
xlLegendPositionBottom = (
    -4107
)  # Enum: XlLegendPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllegendposition
xlLegendPositionCorner = 2  # Enum: XlLegendPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllegendposition
xlLegendPositionCustom = (
    -4161
)  # Enum: XlLegendPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllegendposition
xlLegendPositionLeft = (
    -4131
)  # Enum: XlLegendPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllegendposition
xlLegendPositionRight = (
    -4152
)  # Enum: XlLegendPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllegendposition
xlLegendPositionTop = (
    -4160
)  # Enum: XlLegendPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllegendposition
xlContinuous = 1  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlDash = (
    -4115
)  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlDashDot = 4  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlDashDotDot = 5  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlDot = (
    -4118
)  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlDouble = (
    -4119
)  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlLineStyleNone = (
    -4142
)  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlSlantDashDot = 13  # Enum: XlLineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xllinestyle
xlMarkerStyleAutomatic = (
    -4105
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleCircle = 8  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleDash = (
    -4115
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleDiamond = 2  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleDot = (
    -4118
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleNone = (
    -4142
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStylePicture = (
    -4147
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStylePlus = 9  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleSquare = 1  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleStar = 5  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleTriangle = 3  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlMarkerStyleX = (
    -4168
)  # Enum: XlMarkerStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlmarkerstyle
xlDownward = (
    -4170
)  # Enum: XlOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlorientation
xlHorizontal = (
    -4128
)  # Enum: XlOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlorientation
xlUpward = (
    -4171
)  # Enum: XlOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlorientation
xlVertical = (
    -4166
)  # Enum: XlOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlorientation
xlParentDataLabelOptionsBanner = 1  # Enum: XlParentDataLabelOptions, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlparentdatalabeloptions
xlParentDataLabelOptionsNone = 0  # Enum: XlParentDataLabelOptions, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlparentdatalabeloptions
xlParentDataLabelOptionsOverlapping = 2  # Enum: XlParentDataLabelOptions, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlparentdatalabeloptions
xlPatternAutomatic = (
    -4105
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternChecker = 9  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternCrissCross = 16  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternDown = (
    -4121
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternGray16 = 17  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternGray25 = (
    -4124
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternGray50 = (
    -4125
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternGray75 = (
    -4126
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternGray8 = 18  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternGrid = 15  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternHorizontal = (
    -4128
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternLightDown = 13  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternLightHorizontal = 11  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternLightUp = 14  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternLightVertical = 12  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternLinearGradient = 4000  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternNone = (
    -4142
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternRectangularGradient = 4001  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternSemiGray75 = 10  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternSolid = 1  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternUp = (
    -4162
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPatternVertical = (
    -4166
)  # Enum: XlPattern, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpattern
xlPrinter = 2  # Enum: XlPictureAppearance, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpictureappearance
xlScreen = 1  # Enum: XlPictureAppearance, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpictureappearance
xlCenterPoint = 5  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlInnerCenterPoint = 8  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlInnerClockwisePoint = 7  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlInnerCounterClockwisePoint = 9  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlMidClockwiseRadiusPoint = 4  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlMidCounterClockwiseRadiusPoint = 6  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlOuterCenterPoint = 2  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlOuterClockwisePoint = 3  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlOuterCounterClockwisePoint = 1  # Enum: XlPieSliceIndex, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpiesliceindex
xlHorizontalCoordinate = 1  # Enum: XlPieSliceLocation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpieslicelocation
xlVerticalCoordinate = 2  # Enum: XlPieSliceLocation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpieslicelocation
xlColumnField = 2  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpivotfieldorientation
xlDataField = 4  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpivotfieldorientation
xlHidden = 0  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpivotfieldorientation
xlPageField = 3  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpivotfieldorientation
xlRowField = 1  # Enum: XlPivotFieldOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlpivotfieldorientation
xlContext = (
    -5002
)  # Enum: XlReadingOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlreadingorder
xlLTR = (
    -5003
)  # Enum: XlReadingOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlreadingorder
xlRTL = (
    -5004
)  # Enum: XlReadingOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlreadingorder
xlRegionLabelOptionsBestFitOnly = 1  # Enum: XlRegionLabelOptions, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlregionlabeloptions
xlRegionLabelOptionsNone = 0  # Enum: XlRegionLabelOptions, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlregionlabeloptions
xlRegionLabelOptionsShowAll = 2  # Enum: XlRegionLabelOptions, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlregionlabeloptions
rgbAliceBlue = 16775408  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbAntiqueWhite = 14150650  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbAqua = 16776960  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbAquamarine = 13959039  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbAzure = 16777200  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBeige = 14480885  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBisque = 12903679  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBlack = 0  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBlanchedAlmond = 13495295  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBlue = 16711680  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBlueViolet = 14822282  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBrown = 2763429  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbBurlyWood = 8894686  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbCadetBlue = 10526303  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbChartreuse = 65407  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbCoral = 5275647  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbCornflowerBlue = 15570276  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbCornsilk = 14481663  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbCrimson = 3937500  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkBlue = 9109504  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkCyan = 9145088  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkGoldenrod = 755384  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkGray = 11119017  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkGreen = 25600  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkGrey = 11119017  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkKhaki = 7059389  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkMagenta = 9109643  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkOliveGreen = 3107669  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkOrange = 36095  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkOrchid = 13382297  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkRed = 139  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkSalmon = 8034025  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkSeaGreen = 9419919  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkSlateBlue = 9125192  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkSlateGray = 5197615  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkSlateGrey = 5197615  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkTurquoise = 13749760  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDarkViolet = 13828244  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDeepPink = 9639167  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDeepSkyBlue = 16760576  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDimGray = 6908265  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDimGrey = 6908265  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbDodgerBlue = 16748574  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbFireBrick = 2237106  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbFloralWhite = 15792895  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbForestGreen = 2263842  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbFuchsia = 16711935  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGainsboro = 14474460  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGhostWhite = 16775416  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGold = 55295  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGoldenrod = 2139610  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGray = 8421504  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGreen = 32768  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGreenYellow = 3145645  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbGrey = 8421504  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbHoneydew = 15794160  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbHotPink = 11823615  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbIndianRed = 6053069  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbIndigo = 8519755  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbIvory = 15794175  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbKhaki = 9234160  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLavender = 16443110  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLavenderBlush = 16118015  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLawnGreen = 64636  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLemonChiffon = 13499135  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightBlue = 15128749  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightCoral = 8421616  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightCyan = 9145088  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightGoldenrodYellow = 13826810  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightGray = 13882323  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightGreen = 9498256  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightGrey = 13882323  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightPink = 12695295  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightSalmon = 8036607  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightSeaGreen = 11186720  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightSkyBlue = 16436871  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightSlateGray = 10061943  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightSlateGrey = 10061943  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightSteelBlue = 14599344  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLightYellow = 14745599  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLime = 65280  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLimeGreen = 3329330  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbLinen = 15134970  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMaroon = 128  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumAquamarine = 11206502  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumBlue = 13434880  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumOrchid = 13850042  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumPurple = 14381203  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumSeaGreen = 7451452  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumSlateBlue = 15624315  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumSpringGreen = 10156544  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumTurquoise = 13422920  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMediumVioletRed = 8721863  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMidnightBlue = 7346457  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMintCream = 16449525  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMistyRose = 14804223  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbMoccasin = 11920639  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbNavajoWhite = 11394815  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbNavy = 8388608  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbNavyBlue = 8388608  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbOldLace = 15136253  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbOlive = 32896  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbOliveDrab = 2330219  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbOrange = 42495  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbOrangeRed = 17919  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbOrchid = 14053594  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPaleGoldenrod = 7071982  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPaleGreen = 10025880  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPaleTurquoise = 15658671  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPaleVioletRed = 9662683  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPapayaWhip = 14020607  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPeachPuff = 12180223  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPeru = 4163021  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPink = 13353215  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPlum = 14524637  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPowderBlue = 15130800  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbPurple = 8388736  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbRed = 255  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbRosyBrown = 9408444  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbRoyalBlue = 14772545  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSalmon = 7504122  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSandyBrown = 6333684  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSeaGreen = 5737262  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSeashell = 15660543  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSienna = 2970272  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSilver = 12632256  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSkyBlue = 15453831  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSlateBlue = 13458026  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSlateGray = 9470064  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSlateGrey = 9470064  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSnow = 16448255  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSpringGreen = 8388352  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbSteelBlue = 11829830  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbTan = 9221330  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbTeal = 8421376  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbThistle = 14204888  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbTomato = 4678655  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbTurquoise = 13688896  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbViolet = 15631086  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbWheat = 11788021  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbWhite = 16777215  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbWhiteSmoke = 16119285  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbYellow = 65535  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
rgbYellowGreen = 3329434  # Enum: XlRgbColor, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrgbcolor
xlColumns = 2  # Enum: XlRowCol, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrowcol
xlRows = 1  # Enum: XlRowCol, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlrowcol
xlScaleLinear = (
    -4132
)  # Enum: XlScaleType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlscaletype
xlScaleLogarithmic = (
    -4133
)  # Enum: XlScaleType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlscaletype
xlSeriesColorGradientStyleDiverging = 1  # Enum: XlSeriesColorGradientStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlseriescolorgradientstyle
xlSeriesColorGradientStyleSequential = 0  # Enum: XlSeriesColorGradientStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlseriescolorgradientstyle
xlSeriesNameLevelAll = (
    -1
)  # Enum: XlSeriesNameLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlseriesnamelevel
xlSeriesNameLevelCustom = (
    -2
)  # Enum: XlSeriesNameLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlseriesnamelevel
xlSeriesNameLevelNone = (
    -3
)  # Enum: XlSeriesNameLevel, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlseriesnamelevel
xlSizeIsArea = 1  # Enum: XlSizeRepresents, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlsizerepresents
xlSizeIsWidth = 2  # Enum: XlSizeRepresents, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlsizerepresents
xlTickLabelOrientationAutomatic = (
    -4105
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelorientation
xlTickLabelOrientationDownward = (
    -4170
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelorientation
xlTickLabelOrientationHorizontal = (
    -4128
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelorientation
xlTickLabelOrientationUpward = (
    -4171
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelorientation
xlTickLabelOrientationVertical = (
    -4166
)  # Enum: XlTickLabelOrientation, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelorientation
xlTickLabelPositionHigh = (
    -4127
)  # Enum: XlTickLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelposition
xlTickLabelPositionLow = (
    -4134
)  # Enum: XlTickLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelposition
xlTickLabelPositionNextToAxis = 4  # Enum: XlTickLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelposition
xlTickLabelPositionNone = (
    -4142
)  # Enum: XlTickLabelPosition, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlticklabelposition
xlTickMarkCross = 4  # Enum: XlTickMark, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltickmark
xlTickMarkInside = 2  # Enum: XlTickMark, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltickmark
xlTickMarkNone = (
    -4142
)  # Enum: XlTickMark, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltickmark
xlTickMarkOutside = 3  # Enum: XlTickMark, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltickmark
xlDays = 0  # Enum: XlTimeUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltimeunit
xlMonths = 1  # Enum: XlTimeUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltimeunit
xlYears = 2  # Enum: XlTimeUnit, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltimeunit
xlExponential = 5  # Enum: XlTrendlineType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltrendlinetype
xlLinear = (
    -4132
)  # Enum: XlTrendlineType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltrendlinetype
xlLogarithmic = (
    -4133
)  # Enum: XlTrendlineType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltrendlinetype
xlMovingAvg = 6  # Enum: XlTrendlineType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltrendlinetype
xlPolynomial = 3  # Enum: XlTrendlineType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltrendlinetype
xlPower = 4  # Enum: XlTrendlineType, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xltrendlinetype
xlUnderlineStyleDouble = (
    -4119
)  # Enum: XlUnderlineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlunderlinestyle
xlUnderlineStyleDoubleAccounting = 5  # Enum: XlUnderlineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlunderlinestyle
xlUnderlineStyleNone = (
    -4142
)  # Enum: XlUnderlineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlunderlinestyle
xlUnderlineStyleSingle = 2  # Enum: XlUnderlineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlunderlinestyle
xlUnderlineStyleSingleAccounting = 4  # Enum: XlUnderlineStyle, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlunderlinestyle
xlVAlignBottom = (
    -4107
)  # Enum: XlVAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvalign
xlVAlignCenter = (
    -4108
)  # Enum: XlVAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvalign
xlVAlignDistributed = (
    -4117
)  # Enum: XlVAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvalign
xlVAlignJustify = (
    -4130
)  # Enum: XlVAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvalign
xlVAlignTop = (
    -4160
)  # Enum: XlVAlign, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvalign
xlValueAscending = 1  # Enum: XlValueSortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvaluesortorder
xlValueDescending = 2  # Enum: XlValueSortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvaluesortorder
xlValueNone = 0  # Enum: XlValueSortOrder, https://docs.microsoft.com/en-us/office/vba/api/powerpoint.xlvaluesortorder
