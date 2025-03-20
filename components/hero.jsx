"use client";

import React, { useEffect, useRef } from "react";
import Image from "next/image";
import { Button } from "@/components/ui/button";
import Link from "next/link";

const HeroSection = () => {
  const imageRef = useRef(null);

  useEffect(() => {
    const imageElement = imageRef.current;

    const handleScroll = () => {
      const scrollPosition = window.scrollY;
      const scrollThreshold = 100;

      if (scrollPosition > scrollThreshold) {
        imageElement.classList.add("scrolled");
      } else {
        imageElement.classList.remove("scrolled");
      }
    };

    window.addEventListener("scroll", handleScroll);
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  return (
    <section className="pt-40 pb-20 px-4 bg-white">
      <div className="container mx-auto text-center">
        <h1 className="text-5xl md:text-8xl lg:text-[105px] pb-6 font-black tracking-tighter text-transparent bg-clip-text bg-gradient-to-r from-gray-800 to-gold-600 drop-shadow-lg">
          Rule Your Wealth <br /> with Precision
        </h1>
        <p className="text-xl text-gray-700 mb-8 max-w-2xl mx-auto font-extrabold tracking-widest leading-snug uppercase bg-gradient-to-r from-gold-600 to-gray-600 bg-clip-text text-transparent">
          A ruthless AI arsenal to command, decode, and master your financial empire instantly.
        </p>
        <div className="flex justify-center space-x-4">
          <Link href="/dashboard">
            <Button
              size="lg"
              className="px-8 bg-gold-600 hover:bg-gold-700 text-black font-semibold shadow-md transition-all duration-150"
            >
              Get Started
            </Button>
          </Link>
        </div>
        <div className="hero-image-wrapper mt-5 md:mt-0" ref={imageRef}></div>
      </div>
    </section>
  );
};

export default HeroSection;
