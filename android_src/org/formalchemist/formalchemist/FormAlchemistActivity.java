package org.formalchemist.formalchemist;

import android.os.Bundle;
import android.view.View;
import android.view.ViewGroup;
import android.widget.FrameLayout;
import android.widget.ImageView;

import org.kivy.android.PythonActivity;

public class FormAlchemistActivity extends PythonActivity {
    private View splashOverlay;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        setTheme(R.style.FormAlchemistNativeSplashTheme);
        super.onCreate(savedInstanceState);
        installNativeSplashFade();
    }

    private void installNativeSplashFade() {
        final ViewGroup content = findViewById(android.R.id.content);
        if (content == null) {
            return;
        }

        final FrameLayout overlay = new FrameLayout(this);
        overlay.setClickable(true);
        overlay.setAlpha(1f);

        final ImageView splash = new ImageView(this);
        splash.setImageResource(R.drawable.presplash_native);
        // Use the same full-screen presentation as the window background drawable
        // so the first native frame and the fading overlay match visually.
        splash.setScaleType(ImageView.ScaleType.FIT_XY);
        splash.setAdjustViewBounds(false);

        final FrameLayout.LayoutParams fill = new FrameLayout.LayoutParams(
                ViewGroup.LayoutParams.MATCH_PARENT,
                ViewGroup.LayoutParams.MATCH_PARENT
        );
        overlay.addView(splash, fill);
        content.addView(overlay, fill);
        overlay.bringToFront();
        splashOverlay = overlay;

        content.post(new Runnable() {
            @Override
            public void run() {
                startSplashFade();
            }
        });
    }

    private void startSplashFade() {
        final View overlay = splashOverlay;
        if (overlay == null) {
            return;
        }

        overlay.animate()
                .alpha(0f)
                .setStartDelay(450L)
                .setDuration(1800L)
                .withEndAction(new Runnable() {
                    @Override
                    public void run() {
                        ViewGroup parent = (ViewGroup) overlay.getParent();
                        if (parent != null) {
                            parent.removeView(overlay);
                        }
                        splashOverlay = null;
                    }
                })
                .start();
    }
}
