package org.formalchemist.formalchemist;

import android.graphics.Color;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ImageView;

import org.kivy.android.PythonActivity;

/**
 * Custom activity for Route B startup.
 *
 * Replaces python-for-android's default SDL2 loading layout/icon/text with a
 * plain full-screen black view. The Python side still calls
 * android.loadingscreen.hide_loading_screen() as soon as Kivy is alive.
 */
public class MainActivity extends PythonActivity {
    @Override
    protected View getLoadingScreen() {
        if (mLottieView != null || mImageView != null) {
            return mLottieView != null ? mLottieView : mImageView;
        }

        ImageView view = new ImageView(this);
        view.setLayoutParams(new ViewGroup.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT));
        view.setBackgroundColor(Color.BLACK);
        view.setScaleType(ImageView.ScaleType.CENTER_CROP);

        mImageView = view;
        return mImageView;
    }
}
