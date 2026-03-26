package org.formalchemist.formalchemist;

import android.graphics.Color;
import android.graphics.drawable.ColorDrawable;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.view.ViewGroup;
import android.view.ViewParent;
import android.widget.ImageView;

import org.kivy.android.PythonActivity;

public class MainActivity extends PythonActivity {
    private static final String TAG = "FormAlchemistMainActivity";
    private static final long LOADING_FADE_MS = 260L;

    private boolean loadingFadeStarted = false;

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

        view.setAlpha(1f);
        mImageView = view;
        loadingFadeStarted = false;
        return view;
    }

    @Override
    protected void removeLoadingScreen() {
        if (mImageView == null) {
            super.removeLoadingScreen();
            return;
        }
        if (loadingFadeStarted) {
            Log.v(TAG, "removeLoadingScreen called again while fade is in progress");
            return;
        }

        final ImageView target = mImageView;
        loadingFadeStarted = true;
        Log.v(TAG, "Fading out native loading screen");

        target.post(new Runnable() {
            @Override
            public void run() {
                if (target != mImageView) {
                    loadingFadeStarted = false;
                    return;
                }

                ViewParent parent = target.getParent();
                if (!(parent instanceof ViewGroup)) {
                    Log.v(TAG, "Loading screen parent missing, clearing image view reference");
                    clearTrackedImageView(target);
                    return;
                }

                target.animate()
                        .alpha(0f)
                        .setDuration(LOADING_FADE_MS)
                        .withEndAction(new Runnable() {
                            @Override
                            public void run() {
                                Log.v(TAG, "Native loading screen fade finished");
                                removeTrackedImageView(target);
                            }
                        })
                        .start();
            }
        });
    }

    private void removeTrackedImageView(ImageView target) {
        if (target == null) {
            loadingFadeStarted = false;
            return;
        }

        target.animate().cancel();
        ViewParent parent = target.getParent();
        if (parent instanceof ViewGroup) {
            ((ViewGroup) parent).removeView(target);
        }
        clearTrackedImageView(target);
    }

    private void clearTrackedImageView(ImageView target) {
        target.setImageDrawable(null);
        target.setAlpha(1f);
        if (mImageView == target) {
            mImageView = null;
        }
        loadingFadeStarted = false;
    }
}
