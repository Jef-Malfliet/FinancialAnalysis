package FinancialAnalysis.extensions.java.lang.String;

import manifold.ext.rt.api.Extension;
import manifold.ext.rt.api.This;

import java.lang.String;

@Extension
public class FormulaExtensions {

    public static String inParenthesis(@This String self) {
        return '(' + self + ')';
    }

    public static String negate(@This String self) {
        return '-' + self;
    }

    public static String add(@This String self, String string) {
        return self + '+' + string;
    }

    public static String subtract(@This String self, String string) {
        return self + '-' + string;
    }

    public static String multipliedBy(@This String self, String string) {
        return self + '*' + string;
    }

    public static String dividedBy(@This String self, String string) {
        return self + '/' + string;
    }

    public static String rounded(@This String self, int i) {
        return String.format("ROUND(%s, %d)", self, i);
    }
}