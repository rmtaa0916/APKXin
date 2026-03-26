package org.formalchemist.formalchemist;

import android.graphics.Color;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ImageView;

import org.kivy.android.PythonActivity;

/**
 * Custom activity that replaces the default python-for-android loading UI
 * with a plain black full-screen view, while still registering it through
 * PythonActivity.mImageView so the normal p4a removal path can dismiss it.
 */
public class MainActivity extends PythonActivity {
    private static final String TAG = "FormAlchemistMainActivity";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        Log.v(TAG, "Custom MainActivity onCreate");
        super.onCreate(savedInstanceState);
    }

    @Override
    protected View getLoadingScreen() {
        Log.v(TAG, "Custom getLoadingScreen");

        if (PythonActivity.mLottieView != null || PythonActivity.mImageView != null) {
            return PythonActivity.mLottieView != null
                    ? PythonActivity.mLottieView
                    : PythonActivity.mImageView;
        }

        ImageView view = new ImageView(this);
        view.setLayoutParams(new ViewGroup.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT));
        view.setBackgroundColor(Color.BLACK);
        view.setScaleType(ImageView.ScaleType.CENTER_CROP);
        view.setClickable(false);
        view.setFocusable(false);

        PythonActivity.mLottieView = null;
        PythonActivity.mImageView = view;
        return view;
    }
}
