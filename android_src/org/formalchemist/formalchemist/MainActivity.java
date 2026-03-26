package org.formalchemist.formalchemist;

import android.graphics.Color;
import android.graphics.drawable.ColorDrawable;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ImageView;

import org.kivy.android.PythonActivity;

public class MainActivity extends PythonActivity {
    private static final String TAG = "FormAlchemistMainActivity";

    public interface NewIntentListener extends PythonActivity.NewIntentListener {}
    public interface ActivityResultListener extends PythonActivity.ActivityResultListener {}

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        Log.v(TAG, "Custom MainActivity onCreate");
        super.onCreate(savedInstanceState);
    }

    @Override
    protected View getLoadingScreen() {
        Log.v(TAG, "Custom getLoadingScreen");

        if (mLottieView != null || mImageView != null) {
            return mLottieView != null ? mLottieView : mImageView;
        }

        ImageView view = new ImageView(this);
        view.setLayoutParams(new ViewGroup.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT));
        view.setScaleType(ImageView.ScaleType.CENTER_CROP);
        view.setClickable(true);
        view.setFocusable(true);
        view.setBackgroundColor(Color.BLACK);

        int resId = getResources().getIdentifier("presplash_native", "drawable", getPackageName());

        if (resId != 0) {
            Log.v(TAG, "Using native startup drawable: presplash_native");
            view.setImageResource(resId);
        } else {
            Log.w(TAG, "No native startup drawable found; falling back to black");
            view.setImageDrawable(new ColorDrawable(Color.BLACK));
        }

        mImageView = view;
        return view;
    }
}
