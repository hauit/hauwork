using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    public interface IQuestionCheck
    {
        string CheckAnswer(int questionNumber, Application a, Document d);
    }
    internal class QuestionCheck : IQuestionCheck
    {
        public string CheckAnswer(int questionNumber, Application a, Document d)
        {
            switch (questionNumber)
            {
                case 1: return Cau1(a, d);
                case 2: return Cau2(a, d);
                case 3: return Cau3(a, d);
                case 4: return Cau4(a, d);
                case 5: return Cau5(a, d);
                case 6: return Cau6(a, d);
                case 7: return Cau7(a, d);
                case 8: return Cau8(a, d);
                case 9: return Cau9(a, d);
                case 10: return Cau10(a, d);
                case 11: return Cau11(a, d);
                case 12: return Cau12(a, d);
                case 13: return Cau13(a, d);
                case 14: return Cau14(a, d);
                case 15: return Cau15(a, d);
                case 16: return Cau16(a, d);
                case 17: return Cau17(a, d);
                case 18: return Cau18(a, d);
                case 19: return Cau19(a, d);
                case 20: return Cau20(a, d);
                case 21: return Cau21(a, d);
                case 22: return Cau22(a, d);
                case 23: return Cau23(a, d);
                case 24: return Cau24(a, d);
                case 25: return Cau25(a, d);
                case 26: return Cau26(a, d);
                case 27: return Cau27(a, d);
                case 28: return Cau28(a, d);
                case 29: return Cau29(a, d);
                case 30: return Cau30(a, d);
                case 31: return Cau31(a, d);
                case 32: return Cau32(a, d);
                case 33: return Cau33(a, d);
                case 34: return Cau34(a, d);
                case 35: return Cau35(a, d);
                case 36: return Cau36(a, d);
                case 37: return Cau37(a, d);
                case 38: return Cau38(a, d);
                case 39: return Cau39(a, d);
                case 40: return Cau40(a, d);
                case 41: return Cau41(a, d);
                case 42: return Cau42(a, d);
                case 43: return Cau43(a, d);
                case 44: return Cau44(a, d);
                case 45: return Cau45(a, d);
                case 46: return Cau46(a, d);
                case 47: return Cau47(a, d);
                case 48: return Cau48(a, d);
                case 49: return Cau49(a, d);
                case 50: return Cau50(a, d);
                case 51: return Cau51(a, d);
                case 52: return Cau52(a, d);
                case 53: return Cau53(a, d);
                case 54: return Cau54(a, d);
                case 55: return Cau55(a, d);
                case 56: return Cau56(a, d);
                case 57: return Cau57(a, d);
                case 58: return Cau58(a, d);
                case 59: return Cau59(a, d);
                case 60: return Cau60(a, d);
                case 61: return Cau61(a, d);
                case 62: return Cau62(a, d);
                case 63: return Cau63(a, d);
                case 64: return Cau64(a, d);
                case 65: return Cau65(a, d);
                case 66: return Cau66(a, d);
                case 67: return Cau67(a, d);
                case 68: return Cau68(a, d);
                case 69: return Cau69(a, d);
                case 70: return Cau70(a, d);
                case 71: return Cau71(a, d);
                case 72: return Cau72(a, d);
                case 73: return Cau73(a, d);
                case 74: return Cau74(a, d);
                case 75: return Cau75(a, d);
                case 76: return Cau76(a, d);
                case 77: return Cau77(a, d);
                case 78: return Cau78(a, d);
                case 79: return Cau79(a, d);
                case 80: return Cau80(a, d);
                case 81: return Cau81(a, d);
                case 82: return Cau82(a, d);
                case 83: return Cau83(a, d);
                case 84: return Cau84(a, d);
                case 85: return Cau85(a, d);
                case 86: return Cau86(a, d);
                case 87: return Cau87(a, d);
                case 88: return Cau88(a, d);
                case 89: return Cau89(a, d);
                case 90: return Cau90(a, d);
                case 91: return Cau91(a, d);
                case 92: return Cau92(a, d);
                case 93: return Cau93(a, d);
                case 94: return Cau94(a, d);
                case 95: return Cau95(a, d);
                case 96: return Cau96(a, d);
                case 97: return Cau97(a, d);
                case 98: return Cau98(a, d);
                case 99: return Cau99(a, d);
                case 100: return Cau100(a, d);
                case 101: return Cau101(a, d);
                case 102: return Cau102(a, d);
                case 103: return Cau103(a, d);
                case 104: return Cau104(a, d);
                case 105: return Cau105(a, d);
                case 106: return Cau106(a, d);
                case 107: return Cau107(a, d);
                case 108: return Cau108(a, d);
                case 109: return Cau109(a, d);
                case 110: return Cau110(a, d);
                case 111: return Cau111(a, d);
                case 112: return Cau112(a, d);
                case 113: return Cau113(a, d);
                case 114: return Cau114(a, d);
                case 115: return Cau115(a, d);
                case 116: return Cau116(a, d);
                case 117: return Cau117(a, d);
                case 118: return Cau118(a, d);
                case 119: return Cau119(a, d);
                case 120: return Cau120(a, d);
                case 121: return Cau121(a, d);
                case 122: return Cau122(a, d);
                case 123: return Cau123(a, d);
                case 124: return Cau124(a, d);
                case 125: return Cau125(a, d);
                default:
                    return "Invalid question number";
            }
        }

        private string Cau1(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau2(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau3(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau4(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau5(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau6(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau7(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau8(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau9(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau10(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau11(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau12(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau13(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau14(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau15(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau16(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau17(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau18(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau19(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau20(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau21(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau22(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau23(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau24(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau25(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau26(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau27(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau28(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau29(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau30(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau31(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau32(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau33(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau34(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau35(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau36(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau37(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau38(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau39(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau40(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau41(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau42(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau43(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau44(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau45(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau46(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau47(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau48(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau49(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau50(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau51(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau52(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau53(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau54(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau55(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau56(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau57(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau58(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau59(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau60(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau61(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau62(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau63(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau64(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau65(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau66(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau67(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau68(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau69(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau70(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau71(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau72(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau73(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau74(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau75(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau76(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau77(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau78(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau79(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau80(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau81(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau82(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau83(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau84(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau85(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau86(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau87(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau88(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau89(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau90(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau91(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau92(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau93(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau94(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau95(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau96(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau97(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau98(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau99(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau100(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau101(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau102(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau103(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau104(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau105(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau106(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau107(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau108(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau109(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau110(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau111(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau112(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau113(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau114(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau115(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau116(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau117(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau118(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau119(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau120(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau121(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau122(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau123(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau124(Application a, Document d)
        {
            throw new NotImplementedException();
        }

        private string Cau125(Application a, Document d)
        {
            throw new NotImplementedException();
        }
    }
}
