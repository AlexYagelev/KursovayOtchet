package com.programmerworld.aspose;

import androidx.lifecycle.ViewModel;

public class MainViewModel extends ViewModel {

    private boolean isCheckboxChecked;

    public boolean isCheckboxChecked() {
        return isCheckboxChecked;
    }

    public void setCheckboxChecked(boolean checkboxChecked) {
        isCheckboxChecked = checkboxChecked;
    }

}