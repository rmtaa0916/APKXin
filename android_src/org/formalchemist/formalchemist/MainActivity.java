package org.formalchemist.formalchemist;

import android.graphics.Color;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.view.ViewGroup;
import android.widget.FrameLayout;

import org.kivy.android.PythonActivity;

/**
 * Custom activity that replaces the default python-for-android SDL loading UI
 * with a plain black full-screen view.
 */
public class MainActivity extends PythonActivity {
    private static final String TAG = "FormAlchemistMainActivity";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        Log.v(TAG, "Custom MainActivity onCreate");
        super.onCreate(savedInstanceState);
        try {
            // Re-apply our loading screen immediately in case bootstrap code
            // added its default view before our Python-side hide call runs.
            showLoadingScreen(getLoadingScreen());
        } catch (Exception e) {
            Log.w(TAG, "Unable to re-apply custom loading screen", e);
        }
    }

    @Override
    protected View getLoadingScreen() {
        Log.v(TAG, "Custom getLoadingScreen");

        FrameLayout view = new FrameLayout(this);
        view.setLayoutParams(new ViewGroup.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT));
        view.setBackgroundColor(Color.BLACK);
        view.setClickable(true);
        view.setFocusable(true);
        return view;
    }
}
